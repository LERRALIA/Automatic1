Attribute VB_Name = "Modul20"
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Global DFÜname  As String
Global RCon As Long
Global gbVerbindungstarten As Boolean

Dim bArt_bzw_Ges_Rab As Boolean
Dim dSumBar As Double
Dim dSumUms As Double

Public Function Hat_dieser_Artikel_Markenbezogene_Rabattgrenze(cArtNr As String, dArtRab As Double, dGesrab As Double, dSonderPreis As Double) As Boolean
On Error GoTo LOKAL_ERROR

    Dim rsrs As DAO.Recordset
    Dim sSQL As String
    Dim sRabOk As String
    Dim sMARKE As String
    Dim dRabatt As Double

    Hat_dieser_Artikel_Markenbezogene_Rabattgrenze = False
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cArtNr) = False Then
        Exit Function
    End If
    
    If dArtRab = 0 And dGesrab = 0 And dSonderPreis = 0 Then
        Exit Function
    End If
    
    
    dRabatt = 0
    
    If dArtRab > 0 Then
        dRabatt = dArtRab
    End If
    
    If dGesrab > 0 Then
        dRabatt = dGesrab
    End If
    
    
    'jetzt die Frage, gibt es die Tabelle der Markenbezogenen Rabatthöhen?
    If NewTableSuchenDBKombi("MARKE", gdBase) = False Then
        Exit Function
    End If
    
    Dim gEingetrageneRabgre As Boolean
    
    gEingetrageneRabgre = False
    
    'jetzt die Frage gibt es Markenbezogenen Rabatthöhen?
    sSQL = "Select * from Marke Where rabattgrenze > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        gEingetrageneRabgre = True
    End If
    rsrs.Close
    
    If gEingetrageneRabgre = False Then
        Exit Function
    End If
    
    
    
    'jetzt die Frage ist der Artikel überhaupt rabattfähig
    sRabOk = "J"
    
    sSQL = "Select Rabatt_OK from Artikel Where artnr = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!RABATT_OK) Then
            sRabOk = rsrs!RABATT_OK
        End If
    End If
    rsrs.Close
    
    If sRabOk = "N" Then
        Exit Function
    End If
    
    
    
    
    'jetzt die Frage Welche Marke des Artikels
    sMARKE = ""
    
    Dim sLinie As String
    
    sLinie = ""
    
    sSQL = "Select lpz from Artikel Where artnr = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LPZ) Then
            sLinie = rsrs!LPZ
        End If
    End If
    rsrs.Close
    
    If sLinie = "" Then
        Exit Function
    End If
    
    Dim sMin_Linr As String
    
    sMin_Linr = ermLiefmitkleinstenLEKPR(cArtNr)
    
    
    
    sSQL = "Select Marke from linbez Where lpz = " & sLinie & " and Linr = " & sMin_Linr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!MARKE) Then
            sMARKE = rsrs!MARKE
        End If
    End If
    rsrs.Close
    
    
    If sMARKE = "" Then
        Exit Function
    End If
    
    
    'Welche Rabatthöhe hat diese Marke?
    
    Dim dRabGrenze As Double
    dRabGrenze = 0
    
    sSQL = "Select Rabattgrenze from Marke where marke = '" & sMARKE & "' and Rabattgrenze > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Rabattgrenze) Then
            dRabGrenze = rsrs!Rabattgrenze
        End If
    End If
    rsrs.Close
    
    If dRabGrenze = 0 Then
        Exit Function
    End If
    
    If dRabGrenze >= dRabatt Then
        Exit Function
    Else
    
        Dim sMess As String
        sMess = "Achtung" & vbCrLf
        sMess = sMess & "Bei diesem Artikel(" & cArtNr & ") wird die zulässige Rabattgrenze(" & dRabGrenze & " %) überschritten." & vbCrLf & vbCrLf
        sMess = sMess & "Bitte korrigieren Sie dies!"
        
        
        MsgBox sMess, vbCritical + vbOKOnly, "Winkiss Hinweis:"
    
    
    
    
    End If
    



Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Hat_dieser_Artikel_Markenbezogene_Rabattgrenze"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermZusatztext(cArtNr As String) As String
On Error GoTo LOKAL_ERROR

    Dim cKVKN As String
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim sSQL As String
    Dim cArtikelArt As String
    
    ermZusatztext = ""
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cArtNr) = False Then
        Exit Function
    End If
    
    sSQL = "Select Artnr from Geschwart "
    sSQL = sSQL & "  Where mutterartnr = " & cArtNr
    sSQL = sSQL & " and IMETI = 0 " 'True "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!artnr) Then
            cArtikelArt = rsrs!artnr
        End If
        
        sSQL = "Select KVKPR1 from artikel where artnr = " & cArtikelArt
        Set rsArt = gdBase.OpenRecordset(sSQL)
        If Not rsArt.EOF Then
        
            If Not IsNull(rsArt!KVKPR1) Then
                cKVKN = rsArt!KVKPR1
            End If
            ermZusatztext = "zzgl. " & Format(cKVKN, "#####0.00") & " € Pfand "
        End If
        rsArt.Close
        
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermZusatztext"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function erm_Coupon_ID(cCouponEAN As String) As String
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim sSQL As String
    
    erm_Coupon_ID = ""
    
    If cCouponEAN = "" Then
        Exit Function
    End If
    
    If IsNumeric(cCouponEAN) = False Then
        Exit Function
    End If
    
    sSQL = "Select COUPON_ID from COUPONREGELN "
    sSQL = sSQL & "  Where COUPON_EAN = '" & cCouponEAN & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!COUPON_ID) Then
            erm_Coupon_ID = rsrs!COUPON_ID
        End If
        
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "erm_Coupon_ID"
    Fehler.gsFehlertext = "Im Programmteil Etiketten drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function erm_Coupon_Details(cCouponEAN As String, sWas As String) As String
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim sSQL As String
    
    erm_Coupon_Details = ""
    
    If cCouponEAN = "" Then
        Exit Function
    End If
    
    If IsNumeric(cCouponEAN) = False Then
        Exit Function
    End If
    
    sSQL = "Select " & sWas & " as Ausgabe from COUPONREGELN "
    sSQL = sSQL & "  Where COUPON_EAN = '" & cCouponEAN & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!Ausgabe) Then
            erm_Coupon_Details = rsrs!Ausgabe
        End If
        
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "erm_Coupon_Details"
    Fehler.gsFehlertext = "Im Programmteil Coupon-Details ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Sub StammdatenblattKundeDrucken(cKnr As String, bohneAntrag As Boolean, Optional sOffeneSumme As String)
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    If cKnr = "" Then
        Exit Sub
    End If
    
    loeschNEW "kuntt", gdBase
    CreateTableT2 "KUNTT", gdBase
    
    sSQL = "Insert into KUNTT select "
    sSQL = sSQL & " TEL "
    sSQL = sSQL & ", FAXNR "
    sSQL = sSQL & ", EMAIL "
    sSQL = sSQL & ", MOBILTEL "
    sSQL = sSQL & ", VORNAME "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", NAME "
    sSQL = sSQL & ", STRASSE "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", STADT "
    sSQL = sSQL & ", TITEL "
    sSQL = sSQL & ", FIRMA "
    sSQL = sSQL & ", datum1  "
    sSQL = sSQL & ", GESCHLECHT "
    sSQL = sSQL & ", KUNDKART "
    sSQL = sSQL & ", RABATT  "
    sSQL = sSQL & ", BONUS "
    sSQL = sSQL & ", KURZTEXT1 "
    sSQL = sSQL & ", KURZTEXT2 "
    sSQL = sSQL & ", NOTIZEN "
    
    sSQL = sSQL & ", Merkmal "
    sSQL = sSQL & ", Merkmal2 "
    sSQL = sSQL & ", Anrede "
    sSQL = sSQL & ", Kuerzel "
    sSQL = sSQL & ", Gesperrt "
    sSQL = sSQL & ", Angelegt "
    
    
    sSQL = sSQL & ", '" & gFirma.FirmaName & "' as FIRMANAME "
    sSQL = sSQL & ", '" & gFirma.strasse & "' as FIRMASTRASSE "
    sSQL = sSQL & ", '" & gFirma.Plz & "' as FIRMAPLZ "
    sSQL = sSQL & ", '" & gFirma.Ort & "' as FIRMAORT "

    sSQL = sSQL & " from Kunden where kundnr = " & cKnr
    gdBase.Execute sSQL, dbFailOnError
    
    If sOffeneSumme <> "" Then
        sSQL = "update kuntt set osum = '" & sOffeneSumme & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If bohneAntrag = True Then
        reportbildschirm "dWKL001b", "aWKL13d"
    Else
    
        reportbildschirm "dWKL001b", "aWKL13a"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "StammdatenblattKundeDrucken"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function DS_Noch_Nicht_Unterschrieben(cKnr As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs       As Recordset
    
    DS_Noch_Nicht_Unterschrieben = True
    
    sSQL = "Select * from KUNDEN where kundnr = " & cKnr & " and DS = True"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        DS_Noch_Nicht_Unterschrieben = False
    End If
    rsrs.Close: Set rsrs = Nothing
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "DS_Noch_Nicht_Unterschrieben"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub DatenschutzblattKundeDrucken(cKnr As String, Optional bsofortdruck As Boolean = False)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If cKnr = "" Then
        Exit Sub
    End If
    
        
    If NewTableSuchenDBKombi("DATENSCHUTZ", gdBase) = False Then
    
        loeschNEW "kuntt", gdBase
        CreateTableT2 "KUNTT", gdBase
        
        sSQL = "Insert into KUNTT select "
        sSQL = sSQL & " TEL "
        sSQL = sSQL & ", FAXNR "
        sSQL = sSQL & ", EMAIL "
        sSQL = sSQL & ", MOBILTEL "
        sSQL = sSQL & ", VORNAME "
        sSQL = sSQL & ", KUNDNR "
        sSQL = sSQL & ", NAME "
        sSQL = sSQL & ", STRASSE "
        sSQL = sSQL & ", PLZ "
        sSQL = sSQL & ", STADT "
        sSQL = sSQL & ", TITEL "
        sSQL = sSQL & ", FIRMA "
        sSQL = sSQL & ", datum1  "
        sSQL = sSQL & ", GESCHLECHT "
        sSQL = sSQL & ", KUNDKART "
        sSQL = sSQL & ", RABATT  "
        sSQL = sSQL & ", BONUS "
        sSQL = sSQL & ", KURZTEXT1 "
        sSQL = sSQL & ", KURZTEXT2 "
        sSQL = sSQL & ", NOTIZEN "
        
        sSQL = sSQL & ", Merkmal "
        sSQL = sSQL & ", Merkmal2 "
        sSQL = sSQL & ", Anrede "
        sSQL = sSQL & ", Kuerzel "
        sSQL = sSQL & ", Gesperrt "
        sSQL = sSQL & ", Angelegt "
        
        
        sSQL = sSQL & ", '" & gFirma.FirmaName & "' as FIRMANAME "
        sSQL = sSQL & ", '" & gFirma.strasse & "' as FIRMASTRASSE "
        sSQL = sSQL & ", '" & gFirma.Plz & "' as FIRMAPLZ "
        sSQL = sSQL & ", '" & gFirma.Ort & "' as FIRMAORT "
    
        sSQL = sSQL & " from Kunden where kundnr = " & cKnr
        gdBase.Execute sSQL, dbFailOnError
        
        If bsofortdruck = True Then
            reportbildschirmToPrinter "aWKL13e"
        Else
            reportbildschirm "dWKL001b", "aWKL13e"
        End If
    
    Else
    
        
        Dim rsrs As Recordset
        
        Dim sElement1 As String
        Dim sElement2 As String
        Dim sElement3 As String
        Dim sElement4 As String
        Dim sElement5 As String
        Dim sElement6 As String
        Dim sElement7 As String
        Dim sElement8 As String
        Dim sElement9 As String
        Dim sElement10 As String
        Dim sElement11 As String
        Dim sElement12 As String
        Dim sElement13 As String
        Dim sElement14 As String
        Dim sElement15 As String
        Dim sElement16 As String
        Dim sElement17 As String
        Dim sElement18 As String
        Dim sElement19 As String
        
        Dim bElement0 As Boolean
        Dim bElement1 As Boolean
        Dim bElement2 As Boolean
        Dim bElement3 As Boolean
        Dim bElement4 As Boolean
        Dim bElement5 As Boolean
        Dim bElement6 As Boolean
        Dim bElement7 As Boolean
        Dim bElement8 As Boolean
        
         bElement0 = False
         bElement1 = False
         bElement2 = False
         bElement3 = False
         bElement4 = False
         bElement5 = False
         bElement6 = False
         bElement7 = False
         bElement8 = False
    
        sSQL = "Select * from Datenschutz"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Element1) Then
                sElement1 = rsrs!Element1
            End If
            If Not IsNull(rsrs!Element2) Then
                sElement2 = rsrs!Element2
            End If
            If Not IsNull(rsrs!Element3) Then
                sElement3 = rsrs!Element3
            End If
            If Not IsNull(rsrs!Element4) Then
                sElement4 = rsrs!Element4
            End If
            If Not IsNull(rsrs!Element5) Then
                sElement5 = rsrs!Element5
            End If
            If Not IsNull(rsrs!Element6) Then
                sElement6 = rsrs!Element6
            End If
            If Not IsNull(rsrs!Element7) Then
                sElement7 = rsrs!Element7
            End If
            If Not IsNull(rsrs!Element8) Then
                sElement8 = rsrs!Element8
            End If
            If Not IsNull(rsrs!Element9) Then
                sElement9 = rsrs!Element9
            End If
            If Not IsNull(rsrs!Element10) Then
                sElement10 = rsrs!Element10
            End If
            If Not IsNull(rsrs!Element11) Then
                sElement11 = rsrs!Element11
            End If
            If Not IsNull(rsrs!Element12) Then
                sElement12 = rsrs!Element12
            End If
            
            If Not IsNull(rsrs!Element13) Then
                sElement13 = rsrs!Element13
            End If
            
            If Not IsNull(rsrs!Element14) Then
                sElement14 = rsrs!Element14
            End If
            
            If Not IsNull(rsrs!Element15) Then
                sElement15 = rsrs!Element15
            End If
            
            If Not IsNull(rsrs!Element16) Then
                sElement16 = rsrs!Element16
            End If
            
            If Not IsNull(rsrs!Element17) Then
                sElement17 = rsrs!Element17
            End If
            
            If Not IsNull(rsrs!Element18) Then
                sElement18 = rsrs!Element18
            End If
            
            If Not IsNull(rsrs!Element19) Then
                sElement19 = rsrs!Element19
            End If
            
            
            
            
            
            
            If Not IsNull(rsrs!PflichtName) Then
                If rsrs!PflichtName Then
                    bElement0 = True
                Else
                    bElement0 = False
                End If
            End If
        
            If Not IsNull(rsrs!PflichtvorName) Then
                If rsrs!PflichtvorName Then
                    bElement1 = True
                Else
                    bElement1 = False
                End If
            End If
            
            If Not IsNull(rsrs!Pflichtstadt) Then
                If rsrs!Pflichtstadt Then
                    bElement2 = True
                Else
                    bElement2 = False
                End If
            End If
        
        
        
        
        
            If Not IsNull(rsrs!PflichtPLZ) Then
                If rsrs!PflichtPLZ Then
                    bElement3 = True
                Else
                    bElement3 = False
                End If
            End If
            
            If Not IsNull(rsrs!PflichtSTRASSE) Then
                If rsrs!PflichtSTRASSE Then
                    bElement4 = True
                Else
                    bElement4 = False
                End If
            End If
        
            If Not IsNull(rsrs!PflichtGEBDATUM) Then
                If rsrs!PflichtGEBDATUM Then
                    bElement5 = True
                Else
                    bElement5 = False
                End If
            End If
        
        
        
        
            If Not IsNull(rsrs!PflichtMAIL) Then
                If rsrs!PflichtMAIL Then
                    bElement6 = True
                Else
                    bElement6 = False
                End If
            End If
        
            If Not IsNull(rsrs!PflichtTEL) Then
                If rsrs!PflichtTEL Then
                    bElement7 = True
                Else
                    bElement7 = False
                End If
            End If
        
            If Not IsNull(rsrs!PflichtMOBIL) Then
                If rsrs!PflichtMOBIL Then
                    bElement8 = True
                Else
                    bElement8 = False
                End If
            End If
            
            
            
            
            
            
            
            
            
            
            
            
            
        End If
        rsrs.Close: Set rsrs = Nothing
    
        
        loeschNEW "KUNTT_INDI", gdBase
        CreateTableT2 "KUNTT_INDI", gdBase
        
        sSQL = "Insert into KUNTT_INDI select "
        sSQL = sSQL & " TEL "
        sSQL = sSQL & ", FAXNR "
        sSQL = sSQL & ", EMAIL "
        sSQL = sSQL & ", MOBILTEL "
        sSQL = sSQL & ", VORNAME "
        sSQL = sSQL & ", KUNDNR "
        sSQL = sSQL & ", NAME "
        sSQL = sSQL & ", STRASSE "
        sSQL = sSQL & ", PLZ "
        sSQL = sSQL & ", STADT "
        sSQL = sSQL & ", TITEL "
        sSQL = sSQL & ", FIRMA "
        sSQL = sSQL & ", datum1  "
        sSQL = sSQL & ", GESCHLECHT "
        sSQL = sSQL & ", KUNDKART "
        sSQL = sSQL & ", RABATT  "
        sSQL = sSQL & ", BONUS "
        sSQL = sSQL & ", KURZTEXT1 "
        sSQL = sSQL & ", KURZTEXT2 "
        sSQL = sSQL & ", NOTIZEN "
        
        sSQL = sSQL & ", Merkmal "
        sSQL = sSQL & ", Merkmal2 "
        sSQL = sSQL & ", Anrede "
        sSQL = sSQL & ", Kuerzel "
        sSQL = sSQL & ", Gesperrt "
        sSQL = sSQL & ", Angelegt "
        sSQL = sSQL & ", Filialnr as Filiale "
        
        sSQL = sSQL & ", '" & gFirma.FirmaName & "' as FIRMANAME "
        sSQL = sSQL & ", '" & gFirma.strasse & "' as FIRMASTRASSE "
        sSQL = sSQL & ", '" & gFirma.Plz & "' as FIRMAPLZ "
        sSQL = sSQL & ", '" & gFirma.Ort & "' as FIRMAORT "
        
        sSQL = sSQL & ", '" & sElement1 & "' as ELEMENT1 "
        sSQL = sSQL & ", '" & sElement2 & "' as ELEMENT2 "
        sSQL = sSQL & ", '" & sElement3 & "' as ELEMENT3"
        
        sSQL = sSQL & ", '" & sElement4 & "' as ELEMENT4 "
        sSQL = sSQL & ", '" & sElement5 & "' as ELEMENT5 "
        sSQL = sSQL & ", '" & sElement6 & "' as ELEMENT6 "
        
        sSQL = sSQL & ", '" & sElement7 & "' as ELEMENT7 "
        sSQL = sSQL & ", '" & sElement8 & "' as ELEMENT8 "
        sSQL = sSQL & ", '" & sElement9 & "' as ELEMENT9 "
        
        sSQL = sSQL & ", '" & sElement10 & "' as ELEMENT10 "
        sSQL = sSQL & ", '" & sElement11 & "' as ELEMENT11 "
        sSQL = sSQL & ", '" & sElement12 & "' as ELEMENT12 "
        
        sSQL = sSQL & ", '" & sElement13 & "' as ELEMENT13 "
        sSQL = sSQL & ", '" & sElement14 & "' as ELEMENT14 "
        
        sSQL = sSQL & ", '" & sElement15 & "' as ELEMENT15 "
        sSQL = sSQL & ", '" & sElement16 & "' as ELEMENT16 "
        
        sSQL = sSQL & ", '" & sElement17 & "' as ELEMENT17 "
        sSQL = sSQL & ", '" & sElement18 & "' as ELEMENT18 "
        
        sSQL = sSQL & ", '" & sElement19 & "' as ELEMENT19 "
        
        
        
        
        If bElement0 Then
            sSQL = sSQL & ", true as PflichtNAME  "
        Else
            sSQL = sSQL & ", false as PflichtNAME  "
        End If
        
        If bElement1 Then
            sSQL = sSQL & ", true as PflichtVORNAME  "
        Else
            sSQL = sSQL & ", false as PflichtVORNAME  "
        End If
        
        If bElement2 Then
            sSQL = sSQL & ", true as PflichtSTADT  "
        Else
            sSQL = sSQL & ", false as PflichtSTADT  "
        End If
        
        
        
        If bElement3 Then
            sSQL = sSQL & ", true as PflichtPLZ  "
        Else
            sSQL = sSQL & ", false as PflichtPLZ  "
        End If
        
        If bElement4 Then
            sSQL = sSQL & ", true as PflichtSTRASSE  "
        Else
            sSQL = sSQL & ", false as PflichtSTRASSE  "
        End If
        
        If bElement5 Then
            sSQL = sSQL & ", true as PflichtGEBDATUM  "
        Else
            sSQL = sSQL & ", false as PflichtGEBDATUM  "
        End If
        
        
        
        
        
        If bElement6 Then
            sSQL = sSQL & ", true as PflichtMAIL  "
        Else
            sSQL = sSQL & ", false as PflichtMAIL  "
        End If
        
        If bElement7 Then
            sSQL = sSQL & ", true as PflichtTEL  "
        Else
            sSQL = sSQL & ", false as PflichtTEL  "
        End If
        
        If bElement8 Then
            sSQL = sSQL & ", true as PflichtMOBIL  "
        Else
            sSQL = sSQL & ", false as PflichtMOBIL  "
        End If
        
        
        
        
        
        
        sSQL = sSQL & " from Kunden where kundnr = " & cKnr
        
        
        gdBase.Execute sSQL, dbFailOnError
        
        
        If gbDS_GEB_DRUCKEN = False Then
            sSQL = " Update KUNTT_INDI set datum1 = null"
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        
        
        
        If bsofortdruck = True Then
            If gbDSKLEIN Then
                reportbildschirmToPrinter "aWKL13f"
            Else
                reportbildschirmToPrinter "aWKL13fgro"
            End If
        Else
            If gbDSKLEIN Then
                reportbildschirm "dWKL001b", "aWKL13f"
            Else
                reportbildschirm "dWKL001b", "aWKL13fgro"
            End If
        End If

    End If
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "DatenschutzblattKundeDrucken"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function Welches_Land(cLandkurz As String) As String
    On Error GoTo LOKAL_ERROR

    Welches_Land = ""
    
    Select Case UCase(cLandkurz)
        Case "D"
            Welches_Land = "Deutschland"
        Case "CH"
            Welches_Land = "Schweiz"
        Case "A"
            Welches_Land = "Österreich"
        Case "B"
            Welches_Land = "Belgien"
        Case "DK"
            Welches_Land = "Dänemark"
        Case "F"
            Welches_Land = "Frankreich"
        Case "I"
            Welches_Land = "Italien"
        Case "FL"
            Welches_Land = "Lichtenstein"
        Case "L"
            Welches_Land = "Luxemburg"
        Case "MO"
            Welches_Land = "Monaco"
        Case "NL"
            Welches_Land = "Niederlande"
        Case "PL"
            Welches_Land = "Polen"
        Case "P"
            Welches_Land = "Portugal"
        Case "E"
            Welches_Land = "Spanien"
        Case Else
            Welches_Land = cLandkurz
    End Select
       
       
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Welches_Land"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub zeigImage_In_Picture_Kasse(imgx As Image, PicX As PictureBox, iDiv As Integer)
    On Error GoTo LOKAL_ERROR

    Dim höhe As Integer
    Dim Breite As Integer
    Dim iTeiler As Integer

    If imgx.Width >= imgx.Height Then
        iTeiler = imgx.Width / iDiv
    Else
        iTeiler = imgx.Height / iDiv
    End If
    
    höhe = imgx.Height / iTeiler
    Breite = imgx.Width / iTeiler
    
    imgx.Height = höhe * Screen.TwipsPerPixelX
    imgx.Width = Breite * Screen.TwipsPerPixelY
    
    PicX.Picture = LoadPicture("")
    
    With PicX
        .BorderStyle = 0
        .Width = imgx.Width
        .Height = imgx.Height
        
        ' Wichtig: AutoRedraw = True
        .AutoRedraw = True
        
        ' Bild aus ImageBox übertragen
        .PaintPicture imgx.Picture, 0, 0, _
        imgx.Width, imgx.Height
      
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "zeigImage_In_Picture_Kasse"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub zeig_Kunden_Image_In_Picture_Kasse(imgx As Image, PicX As PictureBox, iDiv As Integer)
    On Error GoTo LOKAL_ERROR

    Dim höhe As Integer
    Dim Breite As Integer
    Dim iTeiler As Integer

    If imgx.Width >= imgx.Height Then
        iTeiler = imgx.Width / iDiv
    Else
        iTeiler = imgx.Height / iDiv
    End If
    
    höhe = imgx.Height / iTeiler
    Breite = imgx.Width / iTeiler
    
    imgx.Height = höhe * Screen.TwipsPerPixelX
    imgx.Width = Breite * Screen.TwipsPerPixelY
    
    PicX.Picture = LoadPicture("")
    
    With PicX
        .BorderStyle = 0
        .Width = imgx.Width
        .Height = imgx.Height
        
        ' Wichtig: AutoRedraw = True
        .AutoRedraw = True
        
        ' Bild aus ImageBox übertragen
        .PaintPicture imgx.Picture, 0, 0, _
        imgx.Width, imgx.Height
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "zeig_Kunden_Image_In_Picture_Kasse"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Drucke_Picklisten()
On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim cDatum      As String
    Dim czeit       As String
    Dim cArtNr      As String
    Dim cBezeich    As String
    Dim cMarke      As String
    Dim cEAN        As String
    Dim cAbgabe     As String
    Dim cFil        As String
    Dim cNPreis     As String
    Dim rsrs        As Recordset
    Dim rsFil       As DAO.Recordset
    Dim iAnzSätze   As Integer
    
    Dim i           As Integer
    Dim lcount      As Long
    
    Dim sSQL        As String

    
    cDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    
    
    
    
    
    
    
    
    sSQL = "Select distinct(Filiale_an) as filiale from PICKLISTE_IN order by Filiale_an"
    
    Set rsFil = gdBase.OpenRecordset(sSQL)
    If Not rsFil.EOF Then
        rsFil.MoveFirst
        Do While Not rsFil.EOF
            
            If Not IsNull(rsFil!FILIALE) Then
                cFil = rsFil!FILIALE
                
                lcount = 0
                
                sSQL = "Select * from PICKLISTE_IN where Filiale_an = " & cFil
    
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveLast
                    iAnzSätze = rsrs.RecordCount
                    ReDim cZeilen(0 To (iAnzSätze * 5) + 5) As String
                    
                    cZeilen(0) = "Diese Artikel tauschen!"
                    cZeilen(1) = "-----------------"
                    cZeilen(2) = "an Filiale: " & cFil
                    cZeilen(3) = "Datum: " & cDatum & "  " & czeit
                    cZeilen(4) = "von: " & gcFilNr
                    cZeilen(5) = vbCrLf
                    
'                    cZeilen(0) = "Diese Artikel tauschen!"
'                    cZeilen(1) = "-----------------"
'                    cZeilen(2) = "an Filiale: " & cFil
'                    cZeilen(3) = "Datum: " & cDatum
'                    cZeilen(4) = "Zeit:  " & czeit
'                    cZeilen(5) = vbCrLf
'
                    
                    rsrs.MoveFirst
                    Do While Not rsrs.EOF
                        
                        If Not IsNull(rsrs!artnr) Then
                            cArtNr = rsrs!artnr
                        End If
                        
                        If Not IsNull(rsrs!kvkpr) Then
                            cNPreis = Format(rsrs!kvkpr, "######.00")
                        End If
                        
                        If Not IsNull(rsrs!BEZEICH) Then
                            cBezeich = rsrs!BEZEICH
                        End If
                        
                        If Not IsNull(rsrs!EAN) Then
                            cEAN = rsrs!EAN
                        End If
                        
                        If Not IsNull(rsrs!MARKE) Then
                            cMarke = rsrs!MARKE
                        End If
                        
                        If Not IsNull(rsrs!Abgabe) Then
                            cAbgabe = rsrs!Abgabe
                        End If
                    
                        cZeilen(6 + lcount) = "Artnr: " & cArtNr & Space(12 - Len(cNPreis)) & cNPreis & " " & gcWaehrung
                        cZeilen(6 + lcount + 1) = cBezeich
                        cZeilen(6 + lcount + 2) = cMarke
                        cZeilen(6 + lcount + 3) = cEAN & " " & cAbgabe & "x (B: " & ermBESTAND(cArtNr) & ")"
                        cZeilen(6 + lcount + 4) = ""
                        
                        lcount = lcount + 5
                        
                    rsrs.MoveNext
                    Loop
                End If
                rsrs.Close: Set rsrs = Nothing
                'Drucke den Beleg
                
                DruckeArbeitszeitBelegWK20d cZeilen(), (iAnzSätze * 5) + 5
                
            End If
            
        rsFil.MoveNext
        Loop
    End If
    rsFil.Close: Set rsFil = Nothing
    
    
    
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Drucke_Picklisten"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SendeDaten2DruckerKUNDENBESTELLUNGWKL20(cKundenNummer As String, cFeldbez1 As String, cFeld1 As String, cFeldbez2 As String, cFeld2 As String, cUeberschrift As String)
    On Error GoTo LOKAL_ERROR
        
    Dim lTask               As Long
    Dim lAnzSatz            As Long
    Dim lAktSatz            As Long
    Dim lcount              As Long
    Dim lAnzZeile           As Long
    Dim lAnzLbSatz          As Long
    Dim lRet                As Long
    
    Dim cLBSatz             As String
    Dim cFeld               As String
    Dim cDaten              As String
    Dim ctmp                As String
    Dim cTmp2               As String
    Dim cMwst               As String
    Dim cText               As String
    Dim aDeviceName         As String
    Dim cEscapeSequenz      As String
    Dim cArtNr              As String
    ReDim cDruckZeile(1 To 1) As String
    
    Dim dGRabatt            As Double
    Dim dGRabattWert        As Double
    Dim dSumme              As Double
    Dim dWert               As Double
    Dim dEuro               As Double
    Dim dMWStVoll           As Double
    Dim dMWStErm            As Double
    Dim dAktZeit            As Double
    Dim dNeuZeit            As Double
    Dim dMWSt               As Double
    
    Dim bAusblenden         As Boolean
    Dim bBonZwang           As Boolean
    Dim iStufe              As Integer
    Dim iLenZeile           As Integer
    Dim iLevel              As Integer
    Dim iAktCopy            As Integer
    Dim iFileNr             As Integer
    Dim bDruckenArtikel     As Boolean
   
    
    bBonZwang = False
    iLevel = 0
    
'    setzedrucker gcBonDrucker
    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz

StartPunkt:
    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    iAktCopy = iAktCopy + 1
    iLevel = 1
    cDaten = ""
    iLenZeile = 32
    dSumme = 0
    dMWStVoll = 0
    dMWStErm = 0
    
    '***********************************************
    'Hier geht's los
    '***********************************************
    
    lAnzSatz = frmWKL20.List1.ListCount
    iLevel = 2
    
    '***********************************************
    'Drucker wird auf BonDrucker geschaltet
    '***********************************************
    
    aDeviceName = gcBonDrucker
    iLevel = 3
    dMWStVoll = 0
    dMWStErm = 0
    
    '***********************************************
    'Drucker ein- und Kundendisplay ausschalten
    '***********************************************
    
    cEscapeSequenz = gcInit
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'ggf. Logo auf Kassenbon bringen
    '***********************************************

    If gcBild <> "" Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Kopfdaten 1.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(0)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 2.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(1)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    '***********************************************
    'Kopfdaten 3.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(4)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        iStufe = 3
    End If
    
    '***********************************************
    'Kopfdaten 4.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(12)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If

    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Default-Text drucken
    '***********************************************
    
'    cDaten = "K U N D E N B E S T E L L U N G"
'    cDaten = "-Kunden---Auslieferung----------"
    cDaten = cUeberschrift
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Artikelpositionen drucken
    '***********************************************
    
    iLevel = 4
    
    dSumme = 0
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        
        cFeld = Mid(cLBSatz, 7, 6)
        cArtNr = cFeld
        
        'hier Überprüfung, befindet sich die Artikelnummer im nichtDruckarray
        
        '3 Arrays auslesen
        
       
        
        Dim cNichtDruckArtnr As String
        Dim iCount As Integer
        bDruckenArtikel = True
        For iCount = 1 To UBound(gcArrArtNr)
            'Artikelnummer
            cNichtDruckArtnr = gcArrArtNr(iCount)
            
            If cFeld = cNichtDruckArtnr Then
                bDruckenArtikel = False 'und somit wird dieser Artikel nicht gedruckt
                Exit For
            End If
            
        Next iCount
        
        If bDruckenArtikel = True Then
        
            If cFeld <> "000000" Then
                '1.Zeile: ArtNr + MWSTKz + ArtBezeich
                cDaten = cFeld & " "
                cFeld = Mid(cLBSatz, 72, 1)
                cDaten = cDaten & cFeld & "  "
                cMwst = cFeld
                cFeld = Mid(cLBSatz, 14, 35)
                cFeld = Trim$(cFeld)
                
                Dim cRestbez As String
                If Len(cFeld) > 17 Then
                    cRestbez = Mid(cFeld, 18, Len(cFeld) - 17)
                    cFeld = Left(cFeld, 17)
                    
                End If
                
                
                cDaten = cDaten & cFeld
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                'neu den Rest Bezeichnung in eine weitere Zeile
                
                If cRestbez <> "" Then
                    
                    cDaten = cRestbez
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
                
                '***********************************************
                'Anzahl, Einzelpreis, Positionspreis drucken
                '***********************************************
                
                If Left(cLBSatz, 1) = "x" Then
                    ctmp = Mid(cLBSatz, 2, 4)
                Else
                    ctmp = Mid(cLBSatz, 1, 5)
                End If
                ctmp = Trim$(ctmp)
                ctmp = ctmp & Space$(6 - Len(ctmp))
                cDaten = ctmp & " x"
                
                If gbRabatt Then
                    ctmp = Mid(cLBSatz, 74, 9)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                Else
                    ctmp = Mid(cLBSatz, 50, 9)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dWert = Val(ctmp)
                End If
                
                ctmp = Format$(dWert, "#####0.00")
                ctmp = Space(11 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                
                ctmp = Mid(cLBSatz, 60, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                dWert = Val(ctmp)
                ctmp = Format$(dWert, "#####0.00")
                dSumme = dSumme + dWert
                ctmp = Space(13 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                
                
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                '***********************************************
                'MWSt-Summe berechnen
                '***********************************************
        
                If cMwst = "V" Then
                    dMWSt = dWert / (100 + gdMWStV)
                    dMWSt = dMWSt * gdMWStV
                    dMWStVoll = dMWStVoll + dMWSt
                ElseIf cMwst = "E" Then
                   dMWSt = dWert / (100 + gdMWStE)
                    dMWSt = dMWSt * gdMWStE
                    dMWStErm = dMWStErm + dMWSt
                Else
                    dMWSt = 0
                End If
            Else
    
                'Zeile mit Zwischensumme drucken
                ctmp = Mid(cLBSatz, 13, Len(cLBSatz) - 13)
                ctmp = Left(Trim(ctmp), 32)
                cDaten = ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
            End If
        End If
    Next lAktSatz
    
    
    
    
    
    
    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    iLevel = 5
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
            
    '***********************************************
    'Endbetrag drucken
    '***********************************************
    
    ctmp = "Endbetrag"
    
    ctmp = Trim$(ctmp)
    ctmp = ctmp & Space$(17 - Len(ctmp))
    ctmp = ctmp & Space$(1) & gcWaehrung

    cDaten = ctmp
    ctmp = Format$(dSumme, "#####0.00")
    ctmp = Space$(11 - Len(ctmp)) & ctmp
    iLevel = 6102
    cDaten = cDaten & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    iLevel = 6103
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************
    cDaten = String$(iLenZeile, "_")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = String$(iLenZeile, "_")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    '***********************************************
    'Zeile Leerzeile drucken
    '***********************************************
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
   
    
    bAusblenden = False
    
    
    ctmp = "Kundenbestellung erstellt von:"
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Bedienername drucken
    '***********************************************
    iLevel = 611
    
    ctmp = gcBediener
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile 'Kassennummer' drucken
    '***********************************************
    
    ctmp = "Kasse: " & gcKasNum
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    If cKundenNummer <> "0" Then
        
        iLevel = 613
        ctmp = "Ihre KundenNr: " & cKundenNummer
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = lookingForKundendaten(cKundenNummer).firma
    
        If ctmp <> "" Then
            If Len(ctmp) > 32 Then
                ctmp = Left(ctmp, 32)
            End If
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
            
        ctmp = lookingForKundendaten(cKundenNummer).titel
        If ctmp <> "" Then
            If Len(ctmp) > 32 Then
                ctmp = Left(ctmp, 32)
            End If
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        ctmp = lookingForKundendaten(cKundenNummer).vorname & " " & lookingForKundendaten(cKundenNummer).nachname
        iLevel = 614
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = lookingForKundendaten(cKundenNummer).strasse
    
        iLevel = 615
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = lookingForKundendaten(cKundenNummer).Plz & " " & lookingForKundendaten(cKundenNummer).Ort
    
        iLevel = 616
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = lookingForKundendaten(cKundenNummer).telefon
    
        iLevel = 616
        ctmp = "Telefon: " & ctmp
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = lookingForKundendaten(cKundenNummer).Mobiltel
    
        iLevel = 616
        ctmp = "Handy: " & ctmp
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
            
        ctmp = lookingForKundendaten(cKundenNummer).Email
    
        iLevel = 616
        
        ctmp = "Email: " & ctmp
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    Else
    
        ctmp = cFeldbez1
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        ctmp = cFeld1
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        ctmp = cFeldbez2
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ctmp = cFeld2
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    iLevel = 615
    
    ctmp = Format$(Date, "DD.MM.YYYY")
    cDaten = ctmp
    ctmp = "0"
    gdBonNr = 0
    iLevel = 6151
    ctmp = gcKasNum & "/" & ctmp
    iLevel = 6152
    ctmp = Space$(8 - Len(ctmp)) & ctmp
    iLevel = 6153
    cDaten = cDaten & Space$(3) & ctmp
    iLevel = 6154
    ctmp = Format$(Now, "HH:MM")
    iLevel = 6155
    cDaten = cDaten & Space$(5) & ctmp
    iLevel = 6156
    KonvertAnsiAscii cDaten
    iLevel = 6157
    cEscapeSequenz = cDaten & vbCrLf
    iLevel = 6158
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    iLevel = 6159
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '***********************************************
    'Zeile Lieferschein bei Kreditverkäufen drucken
    '***********************************************
    
    iLevel = 7
    
    'entfällt bei RETOURE
    '***********************************************
    '1.Zeile Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, gsSTERNZEICH)
'    cDaten = String$(iLenZeile, "*")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Euro drucken
    '***********************************************
    'entfällt bei RETOURE
    '***********************************************
    '2.Zeile Trennstrich drucken
    '***********************************************
    'entfällt bei RETOURE
    '***********************************************
    'Fußzeile 1 drucken
    '***********************************************
    'Fußzeilen
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 2 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iLevel = 10
    End If
    '***********************************************
    'Fußzeile 3 drucken
    '***********************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = ""
    Else
        cDaten = gcBonText(5)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iLevel = 10
    End If
    '***********************************************
    'ein paar Leerzeilen drucken
    '***********************************************
    For lcount = 1 To 9
        If lcount = 9 Then
            cEscapeSequenz = "." & vbCrLf
        Else
            cEscapeSequenz = " " & vbCrLf
        End If
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
    '*************************************************
    '* OpenDrawer3 benutzt die WindowsAPI
    '* OpenDrawer4 geht über das PRINTER-Objekt
    '*************************************************
    '******************************
    'Schublade nur einmal öffnen
    '******************************
    'entfällt bei RETOURE
    iLevel = 12
    
    '******************************
    '* Kassenschublade öffnen
    '******************************

    If gbKBSCHUB Then
        If gbLadeCom Then
            OpenDrawerViaComPortModul20
        Else
            If gbAPI = True Then
                aDeviceName = Printer.DeviceName
                cEscapeSequenz = gcLade
                OpenDrawer aDeviceName, cEscapeSequenz
            End If
        End If
    End If
'
BON_DRUCKEN:
    '**************************************
    '* Bon wird bei RETOURE immer gedruckt
    '**************************************
        If gbAPI = True Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
    
    If iAktCopy = 1 Then
        'Bon-Daten sichern
        If cUeberschrift = "K U N D E N B E S T E L L U N G" Then
            SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False
        End If
    End If

BON_SCHNEIDEN:

    If gbBonDruck Then

        'Kassenbon abschneiden
        If gbAPI = True Then
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcSchneiden
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
        
    End If
    
    iLevel = 11
    
ZWEITER_BON:

    If gb2BONKB = True Then
        If iAktCopy < 2 Then
            GoTo StartPunkt
        End If
    End If

    'entfällt bei RETOURE
    Erase cDruckZeile
    
GUTSCHEIN:
    'entfällt bei RETOURE
ENDE:

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "SendeDaten2DruckerKUNDENBESTELLUNGWKL20"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Public Function ermaktUmsatz(bNurUMSOK As Boolean) As Double
On Error GoTo LOKAL_ERROR
    
Dim sSQL        As String
Dim rsrs       As Recordset

ermaktUmsatz = 0

sSQL = "Select sum(apreis)as maxi from Afcbuch where kasnum = " & gcKasNum

If bNurUMSOK = True Then
    sSQL = sSQL & " and Ums_ok = 'J' "
End If

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        ermaktUmsatz = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermaktUmsatz"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub ProtokolliereRueckGutscheinWK20g(dRueck As Double, lGutschnr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lPos        As Long
    Dim cSatz       As String
    Dim cDatum      As String
    Dim czeit       As String
    Dim cRueck      As String
    Dim cBonNr      As String
    Dim cKasnum     As String
    Dim cGutschnr   As String
    Dim cPfad       As String
    Dim iFileNr     As Integer
    Dim cZeile2     As String
    Dim cSatz1 As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"
    
    cDatum = Format$(Now, "DD.MM.YYYY")                 'Datum der Gutschein-Generierung
    czeit = Format$(Now, "HH:MM:SS")                    'Uhrzeit der Gutschein-Generierung
    cKasnum = gcKasNum                                  'Kassennummer
    cKasnum = Space$(2 - Len(cKasnum)) & cKasnum
    cBonNr = Format$(gdBonNr, "#####0")                 'Bon-Nummer
    cBonNr = Space$(6 - Len(cBonNr)) & cBonNr
    cRueck = Format$(dRueck, "######0.00")              'Was ist davon als Gutschein
    cRueck = Space$(10 - Len(cRueck)) & cRueck
    cGutschnr = Format$(lGutschnr, "#######0")          'Gutscheinnummer
    cGutschnr = Space$(10 - Len(cGutschnr)) & cGutschnr
    
    cSatz1 = "Datum      Uhrzeit   Kasse    Bon    "
    cSatz1 = cSatz1 & "    Wert      GutschNr"
    cSatz1 = cSatz1 & Chr$(13) & Chr$(10)
    
    cSatz = cDatum & " " & czeit & " " & cKasnum & "      " & cBonNr & "  "
    cSatz = cSatz & cRueck & " " & cGutschnr
    cSatz = cSatz & Chr$(13) & Chr$(10)
    cSatz = cSatz & Chr$(13) & Chr$(10)
    
    cGutschnr = Trim$(cGutschnr)
    cRueck = Trim(cRueck)
    gcRueckGutsch = "RüGutsch " & Space$(8 - Len(cGutschnr)) & cGutschnr & " " & gcWaehrung & Space$(11 - Len(cRueck)) & cRueck
    
    iFileNr = FreeFile
    Open cPfad & "GEN_GUT.TXT" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cPfad & "GEN_GUT.TXT"
    End If
    
    Kill cPfad & "GEN_GUT.TXT"
    
    iFileNr = FreeFile
    Open cPfad & "GEN_GUT.TXT" For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz1
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
    
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
        Fehler.gsFormular = "Modul20"
        Fehler.gsFunktion = "ProtokolliereRueckGutscheinWK20g"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
        
        Fehlermeldung1
    End If
End Sub
Public Sub leseBonusBonTexte()
On Error GoTo LOKAL_ERROR

Dim rsrs As Recordset
Dim sSQL As String

gsTextVor = ""
gsTextNach = ""

sSQL = "Select * from BonusBonTexte"
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!TextVor) Then
        gsTextVor = rsrs!TextVor
    End If
    
    If Not IsNull(rsrs!Textnach) Then
        gsTextNach = rsrs!Textnach
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "leseBonusBonTexte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub leseBWWBonTexte()
On Error GoTo LOKAL_ERROR

Dim rsrs As Recordset
Dim sSQL As String

gsTextVor = ""
gsTextNach = ""

gsWWZeichen = ""
gsWWwert = ""
gsWWArt = ""
gsWWSchwellenwert = "0"
gbWWKundBi = False
gsWWBonusArtnr = "0"
gsWWBonusGDAUER = "0"

sSQL = "Select * from BWWBonTexte"
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!TextVor) Then
        gsTextVor = rsrs!TextVor
    End If
    
    If Not IsNull(rsrs!Textnach) Then
        gsTextNach = rsrs!Textnach
    End If
    
    If Not IsNull(rsrs!art) Then
        gsWWArt = rsrs!art
    End If
    
    If Not IsNull(rsrs!SchwellenWert) Then
        gsWWSchwellenwert = rsrs!SchwellenWert
    End If
    
    If Not IsNull(rsrs!BonusArtnr) Then
        gsWWBonusArtnr = rsrs!BonusArtnr
    End If
    
    If Not IsNull(rsrs!GDAUER) Then
        gsWWBonusGDAUER = rsrs!GDAUER
    End If
    
    If Not IsNull(rsrs!Wert) Then
        gsWWwert = rsrs!Wert
    End If
    
    If Not IsNull(rsrs!zeichen) Then
        gsWWZeichen = rsrs!zeichen
    End If
    
    If Not IsNull(rsrs!KUNDBI) Then
        If rsrs!KUNDBI = -1 Then
            gbWWKundBi = False
        Else
            gbWWKundBi = True
        End If
    End If
    
End If
rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "leseBWWBonTexte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub leseBonusArt()
On Error GoTo LOKAL_ERROR

Dim rsrs As Recordset
Dim sSQL As String

giBonusNr = -1

sSQL = "Select * from Bonusart"
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!nr) Then
        giBonusNr = rsrs!nr
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "leseBonusArt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertGarantie(lDat As Long, czeit As String, cKass As String, ibednu1 As Integer, sSERIENNR As String, _
sBemerk As String, lartnr As Long, sBez As String, lKUNDNR As Long, lBonnr As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from Garantie"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!kasnum = cKass
    rsGZ!BEDNU = ibednu1
    
    rsGZ!Seriennr = Trim(sSERIENNR)
    rsGZ!Bemerk = Trim(sBemerk)
    rsGZ!artnr = lartnr
    rsGZ!BEZEICH = sBez
    rsGZ!Kundnr = lKUNDNR
    rsGZ!BELEGNR = lBonnr
    

    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    rsGZ!GEDRUCKT = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "insertGarantie"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function sind_Garatie_daten_zu_drucken(dBonnr As Double) As Boolean
On Error GoTo LOKAL_ERROR

    Dim cSQL                As String
    Dim rsGD                As Recordset
    Dim lMaxNr              As Long
    
    
    Dim lHeute As Long
    lHeute = Fix(Now)
    
    
    
    
    
    
    sind_Garatie_daten_zu_drucken = False
    
    cSQL = "Select Max(lfnr) as maxi  from Garantie "
    
    
    FnOpenrecordset rsGD, cSQL, 1, gdBase
    If Not rsGD.EOF Then
        If Not IsNull(rsGD!maxi) Then
            lMaxNr = rsGD!maxi
        End If
    End If
    rsGD.Close: Set rsGD = Nothing
    
    cSQL = "Select distinct(belegnr)  from Garantie where lfnr = " & lMaxNr
    cSQL = cSQL & " and adate = " & Trim$(Str$(lHeute)) & " "
    
    FnOpenrecordset rsGD, cSQL, 1, gdBase
    If Not rsGD.EOF Then
        If Not IsNull(rsGD!BELEGNR) Then
            If dBonnr = rsGD!BELEGNR Then
                sind_Garatie_daten_zu_drucken = True
            End If
        End If
    End If
    rsGD.Close: Set rsGD = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "sind_Garatie_daten_zu_drucken"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermGarantie_daten(dBonnr As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL                As String
    Dim rsGD                As Recordset
    Dim lHeute              As Long
    Dim lAnzahl             As Long
    Dim lMaxlfnr            As Long
    Dim czeit               As String
    
    
    
    lHeute = Fix(Now)
    lMaxlfnr = 0
    
    'erst max lfnr von dieser Bonnumer suchen, falls es zwei gibt

    cSQL = "Select Top 1 lfnr from Garantie where "
    cSQL = cSQL & " adate = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & " and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & " and BELEGNR = " & dBonnr & " "
    cSQL = cSQL & " order by lfnr desc "
    
    FnOpenrecordset rsGD, cSQL, 1, gdBase
    If Not rsGD.EOF Then
        rsGD.MoveFirst
        If Not IsNull(rsGD!lfnr) Then
            lMaxlfnr = rsGD!lfnr
        End If
    End If
    rsGD.Close: Set rsGD = Nothing
    
    
    
    'mit gedruckt-Flag,kasnum,lheute,dbon alle (eventuell mehrere) rausholen
    
    lAnzahl = 0
    
    cSQL = "Select * from Garantie where "
    cSQL = cSQL & " adate = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & " and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & " and BELEGNR = " & dBonnr & " "
    cSQL = cSQL & " and GEDRUCKT = False "
    
    FnOpenrecordset rsGD, cSQL, 1, gdBase
    If Not rsGD.EOF Then
        rsGD.MoveFirst
        Do While Not rsGD.EOF
        
            lAnzahl = lAnzahl + 1
            
            ReDim Preserve gcArrArtNr(0 To lAnzahl) As String
            If Not IsNull(rsGD!artnr) Then
                gcArrArtNr(lAnzahl) = rsGD!artnr
            Else
                gcArrArtNr(lAnzahl) = ""
            End If
            
            ReDim Preserve gcArrSerienNr(0 To lAnzahl) As String
            If Not IsNull(rsGD!Seriennr) Then
                gcArrSerienNr(lAnzahl) = rsGD!Seriennr
            Else
                gcArrSerienNr(lAnzahl) = ""
            End If
            
            ReDim Preserve gcArrBemerk(0 To lAnzahl) As String
            If Not IsNull(rsGD!Bemerk) Then
                gcArrBemerk(lAnzahl) = rsGD!Bemerk
            Else
                gcArrBemerk(lAnzahl) = ""
            End If
        
            rsGD.MoveNext
        Loop
        
    
    End If
    rsGD.Close: Set rsGD = Nothing
    
    cSQL = "Update Garantie set GEDRUCKT = true where "
    cSQL = cSQL & " adate = " & Trim$(Str$(lHeute)) & " "
    cSQL = cSQL & " and KASNUM = " & gcKasNum & " "
    cSQL = cSQL & " and BELEGNR = " & dBonnr & " "
    cSQL = cSQL & " and GEDRUCKT = False "
    gdBase.Execute cSQL, dbFailOnError
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermGarantie_daten"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ZeigeArtmerk(cART As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset

    ZeigeArtmerk = ""
    cSQL = "Select MERK from ARTMERK where ARTNR = " & cART & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!merk) Then
            ZeigeArtmerk = rsrs!merk
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ZeigeArtmerk"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ZeigeArtKondi(cART As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs1 As DAO.Recordset

    ZeigeArtKondi = ""
    
    cSQL = "Select * from KONDITIONEN where ARTNR = " & cART & " "
    Set rsrs1 = gdBase.OpenRecordset(cSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst

        If Not IsNull(rsrs1!kondi) Then
            ZeigeArtKondi = rsrs1!kondi
        End If

        If Not IsNull(rsrs1!Faktor) Then
            ZeigeArtKondi = ZeigeArtKondi & " + " & rsrs1!Faktor
        End If
    End If
    rsrs1.Close: Set rsrs1 = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ZeigeArtKondi"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ZeigeSTORNOF(cART As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset

    ZeigeSTORNOF = ""
    cSQL = "Select MERK from STORNOF where ARTNR = " & cART & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!merk) Then
            ZeigeSTORNOF = rsrs!merk
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ZeigeSTORNOF"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub InsertProvision()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim cSQL        As String
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim cExtend     As String
    Dim dLiNr       As Double
    Dim dEkpr       As Double
    Dim dWert       As Double
    Dim iFeld       As Integer
    Dim iDbNr       As Integer
    Dim rsrs        As Recordset
    Dim rsKJ        As Recordset
    Dim rsArt       As Recordset
    Dim cArtMWSt    As String
    
    '*** KASSJOUR-Felder ****
    Dim cKJArtNr    As String
    Dim cKJBezeich  As String
    Dim cKJMenge    As String
    Dim cKJAZeit    As String
    Dim cKJKundNr   As String
    Dim cKJFiliale  As String
    Dim cKJKasNum   As String
    Dim cKJLiNr     As String
    Dim cKJLPZ      As String
    Dim cKJAGN      As String
    Dim cKJEAN      As String
    Dim cKJBelegNr  As String
    Dim cUmsOK      As String
    Dim cBonusOk    As String
    Dim cKJMopreis  As String
    
    Dim dKJEkpr     As Double
    Dim dKJVkpr     As Double
    Dim dKJPreis    As Double
    Dim dKJBest1    As Double
    Dim dVkPr       As Double
    Dim dKJPreis2   As Double
    Dim dSpanne     As Double
    Dim lKJADate    As Long
    Dim lKJBediener As Long
    Dim sArtnr      As String
    Dim IAbschluss  As Long
    Dim ierrz       As Integer
    Dim dGeldwert   As Double
    Dim sProvKz     As String
    
    lAnzSatz = frmWKL20.List1.ListCount
    For lAktSatz = 0 To lAnzSatz - 1
        iFeld = 1
        cKJArtNr = ""
        cKJBezeich = ""
        cKJMenge = ""
        dKJPreis = 0
        lKJADate = 0
        cKJAZeit = ""
        lKJBediener = 0
        cKJKundNr = ""
        cKJFiliale = ""
        cKJKasNum = ""
        cKJLiNr = ""
        cKJLPZ = ""
        cKJAGN = ""
        cKJEAN = ""
        dKJEkpr = 0
        dKJVkpr = 0
        cKJBelegNr = ""
        dKJBest1 = 0
        
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)

        'Besonderheiten am Satzende

        'hier Besonders Merkmal - wird in Mopreis kassjour gespeichert
        
        If Len(cLBSatz) > 175 Then
            cKJMopreis = Mid(cLBSatz, 177, 8)
        Else
            cKJMopreis = "0"
        End If

        If Len(cLBSatz) > 157 Then
            cExtend = Mid(cLBSatz, 158, 18)
        Else
            cExtend = ""
        End If
        
        
        
        ctmp = Mid(cLBSatz, 7, 6)
        ctmp = Trim$(ctmp)
        sArtnr = ctmp
        
        sProvKz = Mid(cLBSatz, 6, 1)
        
        '***************************************************
        '* Provisionsartikel darf nicht übernommen werden!
        '***************************************************
        
        If sProvKz = "p" Then
            
            cSQL = "Select * from Artikel where Artnr = " & sArtnr
            FnOpenrecordset rsArt, cSQL, 1, gdBase


            If Not rsArt.EOF Then
                iFeld = 2
                If Not IsNull(rsArt!LPZ) Then
                    cKJLPZ = rsArt!LPZ
                Else
                    cKJLPZ = ""
                End If
                
                iFeld = 3
                If Not IsNull(rsArt!AGN) Then
                    cKJAGN = rsArt!AGN
                Else
                    cKJAGN = ""
                End If
                
                iFeld = 4
                If Not IsNull(rsArt!EAN) Then
                    cKJEAN = rsArt!EAN
                Else
                    cKJEAN = ""
                End If
                
                iFeld = 5
                If Not IsNull(rsArt!ekpr) Then
                    dEkpr = rsArt!ekpr
                Else
                    dEkpr = 0
                End If
                
                iFeld = 6
                If Not IsNull(rsArt!linr) Then
                    dLiNr = rsArt!linr
                Else
                    dLiNr = 0
                End If
                
                iFeld = 7
                If Not IsNull(rsArt!MWST) Then
                    cArtMWSt = rsArt!MWST
                Else
                    cArtMWSt = "V"
                End If
                
                iFeld = 8
                If Not IsNull(rsArt!UMS_OK) Then
                    cUmsOK = rsArt!UMS_OK
                Else
                    cUmsOK = "J"
                End If
                
                iFeld = 9
                
                If Not IsNull(rsArt!BONUS_OK) Then
                    cBonusOk = rsArt!BONUS_OK
                Else
                    cBonusOk = "J"
                End If
                
                If Not IsNull(rsArt!SPANNE) Then
                    dSpanne = rsArt!SPANNE
                Else
                    dSpanne = 0
                End If
                             
            Else
                dEkpr = 0
                dLiNr = 0
            End If
            
            rsArt.Close: Set rsArt = Nothing
            
            
            
            
            If sArtnr = "666666" Then
                cBonusOk = "N"
                cUmsOK = "N"
                cArtMWSt = "O"
            End If
            
            
            ctmp = Mid(cLBSatz, 148, 3)
            ctmp = Trim$(ctmp)
            lKJBediener = Val(ctmp)
            
            
            iFeld = 9
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            cKJMenge = ctmp
            
            iFeld = 10
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dKJPreis = CDbl(ctmp)
            cKJArtNr = sArtnr
            
            iFeld = 12
            ctmp = Mid(cLBSatz, 14, 35)
            ctmp = Trim$(ctmp)
            If cExtend <> "" Then
                If Len(ctmp) > 15 Then
                    ctmp = Left(ctmp, 15)
                End If
                ctmp = ctmp & " @" & cExtend
            End If
            cKJBezeich = ctmp
            
            lKJADate = Fix(Now)
            cKJAZeit = Format$(Now, "HH:MM:SS")
            
            iFeld = 17
            ctmp = frmWKL20.Label2(7).Caption
            ctmp = Trim$(ctmp)
            If Val(ctmp) < 0 Then
                ctmp = "0"
            End If
            cKJKundNr = ctmp
            
    
            cKJBelegNr = gdBonNr
            cKJFiliale = gcFilNr
            
            iFeld = 21
            ctmp = Mid(cLBSatz, 50, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
            iFeld = 22
            ctmp = Mid(cLBSatz, 128, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
           
            If Val(ctmp) = 0 Then
                ctmp = Mid(cLBSatz, 50, 9)
                ctmp = Trim$(ctmp)
                ctmp = fnMoveComma2Point$(ctmp)
                dKJVkpr = Val(ctmp)
            Else
                dKJVkpr = Val(ctmp)
            End If
            
            If dEkpr = 0 Then
                If dSpanne <> 0 Then
                    dEkpr = EKausNettospanneerrechnen(dSpanne, Val(ctmp), cArtMWSt)
                End If
            End If
            
            dKJEkpr = dEkpr
            cKJLiNr = Trim$(Str$(dLiNr))
            
            iFeld = 26
            ctmp = Mid(cLBSatz, 138, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dKJBest1 = Val(ctmp)
            
            cSQL = "Select * from PROVISION where ARTNR = -1"
            FnOpenrecordset rsKJ, cSQL, 1, gdBase
            

            rsKJ.AddNew
            iFeld = 26
            rsKJ!artnr = Val(cKJArtNr)
            iFeld = 27
            rsKJ!BEZEICH = cKJBezeich
            iFeld = 28
            rsKJ!Menge = Val(cKJMenge)
            iFeld = 29
            rsKJ!Preis = dKJPreis
            iFeld = 30
            rsKJ!ADATE = lKJADate
            iFeld = 31
            rsKJ!AZEIT = cKJAZeit
            iFeld = 32
            rsKJ!BEDIENER = lKJBediener
            iFeld = 33
            rsKJ!Kundnr = Val(cKJKundNr)
            iFeld = 34
            rsKJ!FILIALE = Val(cKJFiliale)
            iFeld = 35
            rsKJ!kasnum = Val(gcKasNum)
            iFeld = 36
            rsKJ!linr = Val(cKJLiNr)
            iFeld = 37
            rsKJ!LPZ = Val(cKJLPZ)
            iFeld = 38
            rsKJ!AGN = Val(cKJAGN)
            iFeld = 39
            rsKJ!EAN = cKJEAN
            iFeld = 40
            rsKJ!MWST = cArtMWSt
            iFeld = 41
            rsKJ!ekpr = dKJEkpr
            iFeld = 42
            
            If Trim(cExtend) = "Fleischerbon" Then
                cUmsOK = "N"
            End If
            
            rsKJ!UMS_OK = cUmsOK
            rsKJ!MOPREIS = cKJMopreis

            
            '//Aenderung : Tabelle Kassjour VKPR
            '//wenn ARTIKEL.kvkpr1= 0 Then KASSJOUR.VKPR = KASSJOUR.PREIS
            If dKJVkpr = 0 Then
                dKJPreis2 = rsKJ!Preis
                
                dVkPr = dKJPreis / cKJMenge
                rsKJ!vkpr = dVkPr
            Else
                rsKJ!vkpr = dKJVkpr
            End If
            rsKJ!BELEGNR = gdBonNr
            rsKJ!best1 = dKJBest1
            If gcKreditKarte <> "" Then
                rsKJ!kk_art = gcKreditKarte
            Else
                rsKJ!kk_art = gcZahlMittel
            End If
            rsKJ.Update
            rsKJ.Close
            
        End If
    Next lAktSatz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "InsertProvision"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub InsertXMarkierung()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim cSQL        As String
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim cExtend     As String
    Dim dLiNr       As Double
    Dim dEkpr       As Double
    Dim dWert       As Double
    Dim iFeld       As Integer
    Dim iDbNr       As Integer
    Dim rsrs        As Recordset
    Dim rsKJ        As Recordset
    Dim rsArt       As Recordset
    Dim cArtMWSt    As String
    
    '*** KASSJOUR-Felder ****
    Dim cKJArtNr    As String
    Dim cKJBezeich  As String
    Dim cKJMenge    As String
    Dim cKJAZeit    As String
    Dim cKJKundNr   As String
    Dim cKJFiliale  As String
    Dim cKJKasNum   As String
    Dim cKJLiNr     As String
    Dim cKJLPZ      As String
    Dim cKJAGN      As String
    Dim cKJEAN      As String
    Dim cKJBelegNr  As String
    Dim cUmsOK      As String
    Dim cBonusOk    As String
    Dim cKJMopreis  As String
    
    Dim dKJEkpr     As Double
    Dim dKJVkpr     As Double
    Dim dKJPreis    As Double
    Dim dKJBest1    As Double
    Dim dVkPr       As Double
    Dim dKJPreis2   As Double
    Dim dSpanne     As Double
    Dim lKJADate    As Long
    Dim lKJBediener As Long
    Dim sArtnr      As String
    Dim IAbschluss  As Long
    Dim ierrz       As Integer
    Dim dGeldwert   As Double
    Dim sProvKz     As String
    
    lAnzSatz = frmWKL20.List1.ListCount
    For lAktSatz = 0 To lAnzSatz - 1
        iFeld = 1
        cKJArtNr = ""
        cKJBezeich = ""
        cKJMenge = ""
        dKJPreis = 0
        lKJADate = 0
        cKJAZeit = ""
        lKJBediener = 0
        cKJKundNr = ""
        cKJFiliale = ""
        cKJKasNum = ""
        cKJLiNr = ""
        cKJLPZ = ""
        cKJAGN = ""
        cKJEAN = ""
        dKJEkpr = 0
        dKJVkpr = 0
        cKJBelegNr = ""
        dKJBest1 = 0
        
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)

        'Besonderheiten am Satzende

        'hier Besonders Merkmal - wird in Mopreis kassjour gespeichert
        
        If Len(cLBSatz) > 175 Then
            cKJMopreis = Mid(cLBSatz, 177, 8)
        Else
            cKJMopreis = "0"
        End If

        If Len(cLBSatz) > 157 Then
            cExtend = Mid(cLBSatz, 158, 18)
        Else
            cExtend = ""
        End If
        
        ctmp = Mid(cLBSatz, 7, 6)
        ctmp = Trim$(ctmp)
        sArtnr = ctmp
        
        sProvKz = Mid(cLBSatz, 1, 1)
        
        '***************************************************
        '* Provisionsartikel darf nicht übernommen werden!
        '***************************************************
        
        If sProvKz = "x" Then
            
            cSQL = "Select * from Artikel where Artnr = " & sArtnr
            FnOpenrecordset rsArt, cSQL, 1, gdBase


            If Not rsArt.EOF Then
                iFeld = 2
                If Not IsNull(rsArt!LPZ) Then
                    cKJLPZ = rsArt!LPZ
                Else
                    cKJLPZ = ""
                End If
                
                iFeld = 3
                If Not IsNull(rsArt!AGN) Then
                    cKJAGN = rsArt!AGN
                Else
                    cKJAGN = ""
                End If
                
                iFeld = 4
                If Not IsNull(rsArt!EAN) Then
                    cKJEAN = rsArt!EAN
                Else
                    cKJEAN = ""
                End If
                
                iFeld = 5
                If Not IsNull(rsArt!ekpr) Then
                    dEkpr = rsArt!ekpr
                Else
                    dEkpr = 0
                End If
                
                iFeld = 6
                If Not IsNull(rsArt!linr) Then
                    dLiNr = rsArt!linr
                Else
                    dLiNr = 0
                End If
                
                iFeld = 7
                If Not IsNull(rsArt!MWST) Then
                    cArtMWSt = rsArt!MWST
                Else
                    cArtMWSt = "V"
                End If
                
                iFeld = 8
                If Not IsNull(rsArt!UMS_OK) Then
                    cUmsOK = rsArt!UMS_OK
                Else
                    cUmsOK = "J"
                End If
                
                iFeld = 9
                
                If Not IsNull(rsArt!BONUS_OK) Then
                    cBonusOk = rsArt!BONUS_OK
                Else
                    cBonusOk = "J"
                End If
                
                If Not IsNull(rsArt!SPANNE) Then
                    dSpanne = rsArt!SPANNE
                Else
                    dSpanne = 0
                End If
                             
            Else
                dEkpr = 0
                dLiNr = 0
            End If
            
            rsArt.Close: Set rsArt = Nothing
            
            
            
            
            If sArtnr = "666666" Then
                cBonusOk = "N"
                cUmsOK = "N"
                cArtMWSt = "O"
            End If
            
            
            ctmp = Mid(cLBSatz, 148, 3)
            ctmp = Trim$(ctmp)
            lKJBediener = Val(ctmp)
            
            
            iFeld = 9
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If
            
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            cKJMenge = ctmp
            
            iFeld = 10
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dKJPreis = CDbl(ctmp)
            cKJArtNr = sArtnr
            
            iFeld = 12
            ctmp = Mid(cLBSatz, 14, 35)
            ctmp = Trim$(ctmp)
            If cExtend <> "" Then
                If Len(ctmp) > 15 Then
                    ctmp = Left(ctmp, 15)
                End If
                ctmp = ctmp & " @" & cExtend
            End If
            cKJBezeich = ctmp
            
            lKJADate = Fix(Now)
            cKJAZeit = Format$(Now, "HH:MM:SS")
            
            iFeld = 17
            ctmp = frmWKL20.Label2(7).Caption
            ctmp = Trim$(ctmp)
            If Val(ctmp) < 0 Then
                ctmp = "0"
            End If
            cKJKundNr = ctmp
            
    
            cKJBelegNr = gdBonNr
            cKJFiliale = gcFilNr
            
            iFeld = 21
            ctmp = Mid(cLBSatz, 50, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
            iFeld = 22
            ctmp = Mid(cLBSatz, 128, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
           
            If Val(ctmp) = 0 Then
                ctmp = Mid(cLBSatz, 50, 9)
                ctmp = Trim$(ctmp)
                ctmp = fnMoveComma2Point$(ctmp)
                dKJVkpr = Val(ctmp)
            Else
                dKJVkpr = Val(ctmp)
            End If
            
            If dEkpr = 0 Then
                If dSpanne <> 0 Then
                    dEkpr = EKausNettospanneerrechnen(dSpanne, Val(ctmp), cArtMWSt)
                End If
            End If
            
            dKJEkpr = dEkpr
            cKJLiNr = Trim$(Str$(dLiNr))
            
            iFeld = 26
            ctmp = Mid(cLBSatz, 138, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dKJBest1 = Val(ctmp)
            
            cSQL = "Select * from Markierung where ARTNR = -1"
            FnOpenrecordset rsKJ, cSQL, 1, gdBase
            

            rsKJ.AddNew
            iFeld = 26
            rsKJ!artnr = Val(cKJArtNr)
            iFeld = 27
            rsKJ!BEZEICH = cKJBezeich
            iFeld = 28
            rsKJ!Menge = Val(cKJMenge)
            iFeld = 29
            rsKJ!Preis = dKJPreis
            iFeld = 30
            rsKJ!ADATE = lKJADate
            iFeld = 31
            rsKJ!AZEIT = cKJAZeit
            iFeld = 32
            rsKJ!BEDIENER = lKJBediener
            iFeld = 33
            rsKJ!Kundnr = Val(cKJKundNr)
            iFeld = 34
            rsKJ!FILIALE = Val(cKJFiliale)
            iFeld = 35
            rsKJ!kasnum = Val(gcKasNum)
            iFeld = 36
            rsKJ!linr = Val(cKJLiNr)
            iFeld = 37
            rsKJ!LPZ = Val(cKJLPZ)
            iFeld = 38
            rsKJ!AGN = Val(cKJAGN)
            iFeld = 39
            rsKJ!EAN = cKJEAN
            iFeld = 40
            rsKJ!MWST = cArtMWSt
            iFeld = 41
            rsKJ!ekpr = dKJEkpr
            iFeld = 42
            
            If Trim(cExtend) = "Fleischerbon" Then
                cUmsOK = "N"
            End If
            
            rsKJ!UMS_OK = cUmsOK
            rsKJ!MOPREIS = cKJMopreis

            
            '//Aenderung : Tabelle Kassjour VKPR
            '//wenn ARTIKEL.kvkpr1= 0 Then KASSJOUR.VKPR = KASSJOUR.PREIS
            If dKJVkpr = 0 Then
                dKJPreis2 = rsKJ!Preis
                
                dVkPr = dKJPreis / cKJMenge
                rsKJ!vkpr = dVkPr
            Else
                rsKJ!vkpr = dKJVkpr
            End If
            rsKJ!BELEGNR = gdBonNr
            rsKJ!best1 = dKJBest1
            If gcKreditKarte <> "" Then
                rsKJ!kk_art = gcKreditKarte
            Else
                rsKJ!kk_art = gcZahlMittel
            End If
            rsKJ.Update
            rsKJ.Close
            
        End If
    Next lAktSatz
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "InsertXMarkierung"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function EKausNettospanneerrechnen(dspann As Double, dVerk As Double, cMW As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim Erg1 As Double
    Dim Erg2 As Double
    Dim Erg3 As Double
    Dim Erg4 As Double

    Select Case Trim(cMW)
        Case Is = "V"
            Erg1 = (dVerk * 100) / (100 + gdMWStV)
        Case Is = "E"
            Erg1 = (dVerk * 100) / (100 + gdMWStE)
        Case Else
            Erg1 = (dVerk * 100) / 100
            
    End Select
      
    Erg2 = (Erg1 * dspann) / 100
    Erg3 = Erg1 - Erg2
        
    EKausNettospanneerrechnen = Format(Erg3, "#####0.00")

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "EKausNettospanneerrechnen"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnHoleMaxKundenNr() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lMaxNr As Long
    
    fnHoleMaxKundenNr = -1
    
    cSQL = "Select max(KUNDNR) from KUNDEN"
    If gbFilNr Then
        cSQL = cSQL & " where KUNDNR > " & gcFilNr & "00000 "
        cSQL = cSQL & " and KUNDNR <= " & gcFilNr & "99999 "
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs.Fields(0)) Then
            lMaxNr = rsrs.Fields(0)
        Else
            lMaxNr = 0
        End If
    Else
        lMaxNr = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    If gbFilNr And lMaxNr = 0 Then
        lMaxNr = Val(gcFilNr & "00000")
    End If
    
    lMaxNr = lMaxNr + 1
    
    fnHoleMaxKundenNr = lMaxNr
    
    
    
    
    If gbFilNr Then
        If lMaxNr <= Val(gcFilNr & "00000") Or lMaxNr > Val(gcFilNr & "99999") Then
            'suche eine nummer im Bereich
            
            cSQL = "Select Top 1 Kundnr + 1 as Kundnr2 "
            cSQL = cSQL & " from kunden t1 where t1.Kundnr + 1 not in  "
            cSQL = cSQL & " (SELECT  Kundnr  FROM kunden t2 where t2.Kundnr = t1.Kundnr + 1) "
            cSQL = cSQL & " and t1.Kundnr between " & gcFilNr & "00000 " & " and " & gcFilNr & "99999 "
            cSQL = cSQL & " and t1.Kundnr + 1 not in (SELECT  Kundnr  FROM kassjour t3 where t3.Kundnr = t1.Kundnr + 1) "
            cSQL = cSQL & " order by Kundnr asc "
        
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!Kundnr2) Then
                    lMaxNr = rsrs!Kundnr2
                End If
            End If
            rsrs.Close
    
            
        End If
    End If
    
    fnHoleMaxKundenNr = lMaxNr
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "fnHoleMaxKundenNr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
    
End Function
Public Function fnIsKundenNrfrei(lKUNDNR As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    fnIsKundenNrfrei = False
    
    cSQL = "Select * from KUNDEN where kundnr = " & lKUNDNR
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        fnIsKundenNrfrei = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "fnIsKundenNrfrei"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub LeseBankLeitZahlWKL20(cKundnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    cSQL = "Select * from BANKEN where BLZ = '" & gECKarte.BLZ & "' "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BankName) Then
            gECKarte.BankName = rsrs!BankName
        Else
            gECKarte.BankName = ""
        End If
        
        If Not IsNull(rsrs!Ort) Then
            gECKarte.BankOrt = rsrs!Ort
        Else
            gECKarte.BankOrt = ""
        End If
        
    Else
        gECKarte.BankName = ""
        gECKarte.BankOrt = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    gECKarte.KontoInhaber = ""
    
    cSQL = "Select * from KUNDEN where ECIDENT = '" & gECKarte.BLZ & gECKarte.Konto1 & "' "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    If rsrs.EOF Then
        If cKundnr <> "0" Then
            cSQL = "Select * from KUNDEN where KUNDNR = " & cKundnr 'Label2(7).Caption
            FnOpenrecordset rsrs, cSQL, 1, gdBase
        End If
    End If
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!vorname) Then
            gECKarte.KontoInhaber = rsrs!vorname
        Else
            gECKarte.KontoInhaber = ""
        End If
        If Not IsNull(rsrs!name) Then
            gECKarte.KontoInhaber = gECKarte.KontoInhaber & " " & rsrs!name
        Else
            gECKarte.KontoInhaber = gECKarte.KontoInhaber
        End If
        
        'jetzt optional nach Hakensetzung
        'Lucks
        If gbNachKBbeiEC Then
            If cKundnr = "0" Then
                If Not IsNull(rsrs!Kundnr) Then
                    gckundnr = rsrs!Kundnr
                End If
            End If
        End If
        
        rsrs.Edit
        rsrs!Status = "E"
        rsrs!SYNStatus = "E"
        rsrs!ECIDENT = gECKarte.BLZ & gECKarte.Konto1
        rsrs.Update
        
    Else
        gECKarte.KontoInhaber = "unbekannt"
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "LeseBankLeitZahlWKL20"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeereDatenECKarteWKL20()
    On Error GoTo LOKAL_ERROR
    
    gECKarte.Datenstrom = ""
    gECKarte.ECSpur1 = ""
    gECKarte.ECSpur2 = ""
    gECKarte.BLZ = ""
    gECKarte.BankName = ""
    gECKarte.BankOrt = ""
    gECKarte.Konto1 = ""
    gECKarte.Konto2 = ""
    gECKarte.jahr = ""
    gECKarte.Monat = ""
    gECKarte.KontoInhaber = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "LeereDatenECKarteWKL20"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub KompressDTAWKL57()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
   
    cSQL = "Delete from DTA"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "TEMPXXXX", gdBase

    cSQL = "Select * into TEMPXXXX from DTA"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "DTA", gdBase

    cSQL = "Select * into DTA from TEMPXXXX"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "KompressDTAWKL57"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub fillecbo(sKurz As String, cbox As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cbox.Clear
    
    If sKurz <> "" Then
    
        cSQL = "Select distinct(vorname) from KUNDEN "
    
        If Len(sKurz) > 5 Then
            cSQL = cSQL & " where name like '" & sKurz & "*' "
        ElseIf Len(sKurz) <= 5 Then
            cSQL = cSQL & " where kuerzel like '" & sKurz & "*' "
        End If
        
        cSQL = cSQL & "  order by vorname "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                If Not IsNull(rsrs!vorname) Then
                    cbox.AddItem rsrs!vorname
                End If
                rsrs.MoveNext
            Loop
    
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
    End If
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "fillecbo"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub sicherdta()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim lWert As Long
    Dim cdatei As String
    Dim ctmp As String
    
    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "DTA" & ctmp & Format$(TimeValue(Now), "HH:MM:SS")
    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    
    loeschNEW cdatei, gdBase
    
    sSQL = "Select * into " & cdatei & " from DTA "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "sicherdta"
    Fehler.gsFehlertext = "Im Programmteil DTA ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub updateafcstat(sSpalte As String, dWert As Double, sKasnum As String)
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim rsrs As Recordset
Dim lDatum As Long

lDatum = Fix(Now)

cSQL = "Select adate,kasnum," & sSpalte & " as Auswahl from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & sKasnum
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    rsrs.Edit
Else
    rsrs.AddNew
    rsrs!ADATE = lDatum
    rsrs!kasnum = sKasnum
End If

If Not IsNull(rsrs!auswahl) Then
    rsrs!auswahl = rsrs!auswahl + dWert
Else
    rsrs!auswahl = dWert
End If
rsrs.Update
rsrs.Close: Set rsrs = Nothing



Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "updateafcstat"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Function ermLinrInZeitE() As Long
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermLinrInZeitE = 0
    
    If Not NewTableSuchenDBKombi("ZEITE", gdBase) Then
        Exit Function
    End If
    

    sSQL = "Select zeitLINR from zeite "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!zeitLINR) Then
            ermLinrInZeitE = rsrs!zeitLINR
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
            
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermLinrInZeitE"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermartnrausLIBESNR(cLiBesNr As String, lLinr As Long) As String
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermartnrausLIBESNR = ""

sSQL = "Select artnr from artikel where LIBESNR = '" & cLiBesNr & "'"
If lLinr > 0 Then
    sSQL = sSQL & " and linr = " & lLinr
End If
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then

    rsrs.MoveFirst
    If Not IsNull(rsrs!artnr) Then
        ermartnrausLIBESNR = rsrs!artnr
    End If

End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    
    Fehler.gsFunktion = "ermartnrausLIBESNR"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermPREIS(cART As String, cPreistyp As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    ermPREIS = 0
    
    Select Case cPreistyp
        Case "LEKPR"
            sSQL = "Select * from Artlief where artnr = " & cART
        Case "LVKPR"
            sSQL = "Select * from Artikel where artnr = " & cART
        
    End Select
    
    
    
    FnOpenrecordset rsrs, sSQL, 1, gdBase
    
    If Not rsrs.EOF Then
    
        Select Case cPreistyp
            Case "LEKPR"
                If Not IsNull(rsrs!lekpr) Then
                    ermPREIS = rsrs!lekpr
                End If
            Case "LVKPR"
                If Not IsNull(rsrs!vkpr) Then
                    ermPREIS = rsrs!vkpr
                End If
        End Select
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermPREIS"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermPREISKZ(cKundnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    ermPREISKZ = "0"
    
    sSQL = "Select PREISKZ from Kunden where Kundnr = " & cKundnr
    FnOpenrecordset rsrs, sSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!PREISKZ) Then
            ermPREISKZ = rsrs!PREISKZ
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermPREISKZ"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub HoleUnterbrochenenBonWK20b(cLfdNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsBonPause As Recordset
    Dim rsKd As Recordset
    
    Dim cBedNr As String
    Dim cZSum As String
    Dim cGRabatt As String
    Dim cKdName As String
    Dim cKdnr As String
    Dim cLbText As String
    Dim cPreisKz As String
    Dim sArtnr As String
    
    cSQL = "Select * from BONPAUSE where LFDNR = " & cLfdNr & " order by LBZEILE "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BEDNR) Then
            cBedNr = rsrs!BEDNR
        Else
            cBedNr = ""
        End If
        If Not IsNull(rsrs!ZSUM) Then
            cZSum = rsrs!ZSUM
        Else
            cZSum = "0,00"
        End If
        If Not IsNull(rsrs!GRABATT) Then
            cGRabatt = rsrs!GRABATT
        Else
            cGRabatt = "0,00"
        End If
        If Not IsNull(rsrs!KdName) Then
            cKdName = rsrs!KdName
        Else
            cKdName = ""
        End If
        If Not IsNull(rsrs!KdNr) Then
            cKdnr = rsrs!KdNr
        Else
            cKdnr = ""
        End If
        
        If cKdnr <> "" Then
            cSQL = "Select * from Kunden where Kundnr = " & Val(cKdnr)
            Set rsKd = gdBase.OpenRecordset(cSQL)
            
            If Not rsKd.EOF Then
                If Not IsNull(rsKd!PREISKZ) Then
                    cPreisKz = rsKd!PREISKZ
                Else
                    cPreisKz = "0"
                End If
                frmWKL20!Label8(3).Caption = cPreisKz
                giPreisKz = Val(cPreisKz)
            End If
            rsKd.Close
        End If
        
        frmWKL20!Label2(6).Caption = cZSum
        If cGRabatt <> "0,00" Then
            frmWKL20!Label2(3).Caption = cGRabatt
            frmWKL20!Label2(3).Visible = True
            frmWKL20!Label1(3).Visible = True
        Else
            frmWKL20!Label2(4).Visible = False
            frmWKL20!Label2(3).Visible = False
            frmWKL20!Label1(3).Visible = False
        End If
        
        If cKdName <> "" Or cKdnr <> "" Then
            frmWKL20!Label1(7).Caption = cKdName
            frmWKL20!Label2(7).Caption = cKdnr
            frmWKL20!Label1(7).Visible = True
            frmWKL20!Label1(19).Visible = True
            frmWKL20!Label2(7).Visible = True
        Else
            frmWKL20!Label1(7).Visible = False
            frmWKL20!Label1(19).Visible = False
            frmWKL20!Label2(7).Visible = False
        End If
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!lbtext) Then
                cLbText = rsrs!lbtext
                
                
                    sArtnr = Mid(cLbText, 7, 6)
                    sArtnr = Trim$(sArtnr)
                    
                    If sArtnr = "666666" Then
                        If gbGutscheinBeiVKversteuern = True Then
                            Mid(cLbText, 72, 1) = "V"
                        End If
                        
                        'auch das ausgabedatum muss angepasst werden
                        
                    End If
                    
                If giFarbebeiPark > 0 And giFarbebeiPark < 20 Then
                    
                    cSQL = " Select count(*)  from bonpause where trim(Mid(lbtext,7,6)) = '" & sArtnr & "'"
                    Set rsBonPause = gdBase.OpenRecordset(cSQL)
                    
                    If Not rsBonPause.EOF Then
                        
                        If rsBonPause.RecordCount = 1 Then
                    
                            cSQL = "Update ARTIKEL set AWM = '0'  where Artnr = " & sArtnr
                            gdBase.Execute cSQL, dbFailOnError
                        
                        End If
                    
                    End If
                    
                    rsBonPause.Close
                    
                End If
                
                
                
                
'                cLbText = RTrim$(cLbText) 'hab ich für die zwischensumme rausgenommen
                frmWKL20!List1.AddItem cLbText
                frmWKL20!List3.Nodes.Add Text:=Left(cLbText, 68)
                farbelist3
            End If
            rsrs.MoveNext
        Loop
        
        cSQL = "Delete from BONPAUSE where LFDNR = " & cLfdNr & " "
        gdBase.Execute cSQL, dbFailOnError
        
        'nur Anzeige
        HoleNeueBonNrWKL20
        
        
        frmWKL20!Label18.Caption = gdBonNr
        frmWKL20!Label18.Refresh
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "HoleUnterbrochenenBonWK20b"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub HoleUnterbrochenenBonWK20b_einzelneArtikel(cLfdNr As String, sLBZeileArr() As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsBonPause As Recordset
    Dim rsKd As Recordset
    
    Dim cBedNr As String
    Dim cZSum As String
    Dim cGRabatt As String
    Dim cKdName As String
    Dim cKdnr As String
    Dim cLbText As String
    Dim cPreisKz As String
    Dim i As Integer
    Dim sArtnr As String
    
    cSQL = "Select * from BONPAUSE where LFDNR = " & cLfdNr & " "
    
    cSQL = cSQL & " and ("
    For i = 1 To UBound(sLBZeileArr)
    
    
    
        cSQL = cSQL & " LBZEILE = " & sLBZeileArr(i)
        cSQL = cSQL & " or "
                    
                    
                
    Next i
    
    cSQL = Mid(cSQL, 1, Len(cSQL) - 3)
    
    cSQL = cSQL & " ) "
    
    
    cSQL = cSQL & " order by LBZEILE "
'    MsgBox cSQL
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BEDNR) Then
            cBedNr = rsrs!BEDNR
        Else
            cBedNr = ""
        End If
        If Not IsNull(rsrs!ZSUM) Then
            cZSum = rsrs!ZSUM
        Else
            cZSum = "0,00"
        End If
        If Not IsNull(rsrs!GRABATT) Then
            cGRabatt = rsrs!GRABATT
        Else
            cGRabatt = "0,00"
        End If
        If Not IsNull(rsrs!KdName) Then
            cKdName = rsrs!KdName
        Else
            cKdName = ""
        End If
        If Not IsNull(rsrs!KdNr) Then
            cKdnr = rsrs!KdNr
        Else
            cKdnr = ""
        End If
        
        If cKdnr <> "" Then
            cSQL = "Select * from Kunden where Kundnr = " & Val(cKdnr)
            Set rsKd = gdBase.OpenRecordset(cSQL)
            
            If Not rsKd.EOF Then
                If Not IsNull(rsKd!PREISKZ) Then
                    cPreisKz = rsKd!PREISKZ
                Else
                    cPreisKz = "0"
                End If
                frmWKL20!Label8(3).Caption = cPreisKz
                giPreisKz = Val(cPreisKz)
            End If
            rsKd.Close
        End If
        
        frmWKL20!Label2(6).Caption = cZSum
        If cGRabatt <> "0,00" Then
            frmWKL20!Label2(3).Caption = cGRabatt
            frmWKL20!Label2(3).Visible = True
            frmWKL20!Label1(3).Visible = True
        Else
            frmWKL20!Label2(4).Visible = False
            frmWKL20!Label2(3).Visible = False
            frmWKL20!Label1(3).Visible = False
        End If
        
        If cKdName <> "" Or cKdnr <> "" Then
            frmWKL20!Label1(7).Caption = cKdName
            frmWKL20!Label2(7).Caption = cKdnr
            frmWKL20!Label1(7).Visible = True
            frmWKL20!Label1(19).Visible = True
            frmWKL20!Label2(7).Visible = True
        Else
            frmWKL20!Label1(7).Visible = False
            frmWKL20!Label1(19).Visible = False
            frmWKL20!Label2(7).Visible = False
        End If
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!lbtext) Then
                cLbText = rsrs!lbtext
                
                If giFarbebeiPark > 0 And giFarbebeiPark < 20 Then
                
                    sArtnr = Mid(cLbText, 7, 6)
                    sArtnr = Trim$(sArtnr)
                    
                    cSQL = " Select count(*)  from bonpause where trim(Mid(lbtext,7,6)) = '" & sArtnr & "'"
                    Set rsBonPause = gdBase.OpenRecordset(cSQL)
                    
                    If Not rsBonPause.EOF Then
                        
                        If rsBonPause.RecordCount = 1 Then
                    
                            cSQL = "Update ARTIKEL set AWM = '0'  where Artnr = " & sArtnr
                            gdBase.Execute cSQL, dbFailOnError
                        
                        End If
                    
                    End If
                    
                    rsBonPause.Close
                
                    
                    
                End If
                
                
'                cLbText = RTrim$(cLbText) 'hab ich für die zwischensumme rausgenommen
                frmWKL20!List1.AddItem cLbText
                frmWKL20!List3.Nodes.Add Text:=Left(cLbText, 68)
                farbelist3
            End If
            rsrs.MoveNext
        Loop
        
        cSQL = "Delete from BONPAUSE where LFDNR = " & cLfdNr & " "
        
        cSQL = cSQL & " and ("
        For i = 1 To UBound(sLBZeileArr)
        
        
        
            cSQL = cSQL & " LBZEILE = " & sLBZeileArr(i)
            cSQL = cSQL & " or "
                        
                        
                    
        Next i
        
        cSQL = Mid(cSQL, 1, Len(cSQL) - 3)
        
        cSQL = cSQL & " ) "
        
        
        
        
        gdBase.Execute cSQL, dbFailOnError
        
        'nur Anzeige
        HoleNeueBonNrWKL20
        
        frmWKL20!Label18.Caption = gdBonNr
        frmWKL20!Label18.Refresh
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "HoleUnterbrochenenBonWK20b_einzelneArtikel"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub HoleUnterbrochenenBonWK20j_ARTAUSWAHL(cLfdNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsKd As Recordset
    
    Dim cBedNr As String
    Dim cZSum As String
    Dim cGRabatt As String
    Dim cKdName As String
    Dim cKdnr As String
    Dim cLbText As String
    Dim cPreisKz As String
    
    cSQL = "Select * from ARTAUSWAHL where LFDNR = " & cLfdNr & " order by LBZEILE "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!BEDNR) Then
            cBedNr = rsrs!BEDNR
        Else
            cBedNr = ""
        End If
        If Not IsNull(rsrs!ZSUM) Then
            cZSum = rsrs!ZSUM
        Else
            cZSum = "0,00"
        End If
        If Not IsNull(rsrs!GRABATT) Then
            cGRabatt = rsrs!GRABATT
        Else
            cGRabatt = "0,00"
        End If
        If Not IsNull(rsrs!KdName) Then
            cKdName = rsrs!KdName
        Else
            cKdName = ""
        End If
        If Not IsNull(rsrs!KdNr) Then
            cKdnr = rsrs!KdNr
        Else
            cKdnr = ""
        End If
        
        If cKdnr <> "" Then
            cSQL = "Select * from Kunden where Kundnr = " & Val(cKdnr)
            Set rsKd = gdBase.OpenRecordset(cSQL)
            
            If Not rsKd.EOF Then
                If Not IsNull(rsKd!PREISKZ) Then
                    cPreisKz = rsKd!PREISKZ
                Else
                    cPreisKz = "0"
                End If
                frmWKL20!Label8(3).Caption = cPreisKz
                giPreisKz = Val(cPreisKz)
            End If
            rsKd.Close
        End If
        
        frmWKL20!Label2(6).Caption = cZSum
        If cGRabatt <> "0,00" Then
            frmWKL20!Label2(3).Caption = cGRabatt
            frmWKL20!Label2(3).Visible = True
            frmWKL20!Label1(3).Visible = True
        Else
            frmWKL20!Label2(4).Visible = False
            frmWKL20!Label2(3).Visible = False
            frmWKL20!Label1(3).Visible = False
        End If
        
        If cKdName <> "" Or cKdnr <> "" Then
            frmWKL20!Label1(7).Caption = cKdName
            frmWKL20!Label2(7).Caption = cKdnr
            frmWKL20!Label1(7).Visible = True
            frmWKL20!Label1(19).Visible = True
            frmWKL20!Label2(7).Visible = True
        Else
            frmWKL20!Label1(7).Visible = False
            frmWKL20!Label1(19).Visible = False
            frmWKL20!Label2(7).Visible = False
        End If
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!lbtext) Then
                cLbText = rsrs!lbtext
                cLbText = RTrim$(cLbText)
                frmWKL20!List1.AddItem cLbText
                frmWKL20!List3.Nodes.Add Text:=Left(cLbText, 68)
                farbelist3
            End If
            rsrs.MoveNext
        Loop
        
        'nur Anzeige
        HoleNeueBonNrWKL20
        
        frmWKL20!Label18.Caption = gdBonNr
        frmWKL20!Label18.Refresh
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "HoleUnterbrochenenBonWK20j_ARTAUSWAHL"
    Fehler.gsFehlertext = "Im Programmteil offene Kassenvorgänge ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function CheckofP() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim cLBSatz As String
    Dim cArtNr As String
    
    CheckofP = False

    For i = 1 To frmWKL20.List3.Nodes.Count
        cLBSatz = frmWKL20.List3.Nodes(i).Text
        cArtNr = Mid(cLBSatz, 6, 1)
        If cArtNr = "p" Then
            CheckofP = True
            Exit For
        End If
    Next i
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "CheckofP"
    Fehler.gsFehlertext = "Im Programmteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function CheckofX() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim cLBSatz As String
    Dim cArtNr As String
    
    CheckofX = False

    For i = 1 To frmWKL20.List3.Nodes.Count
        cLBSatz = frmWKL20.List3.Nodes(i).Text
        cArtNr = Mid(cLBSatz, 1, 1)
        If cArtNr = "x" Then
            CheckofX = True
            Exit For
        End If
    Next i
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "CheckofX"
    Fehler.gsFehlertext = "Im Programmteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub farbelist3()
    On Error GoTo LOKAL_ERROR
    
    
    Dim byAWM As Byte
    Dim i As Integer
    Dim j As Integer
    Dim cLBSatz As String
    Dim cArtNr As String
    
    
    If gsFARBKASSE = "1" Then
        For i = 1 To frmWKL20.List3.Nodes.Count
    
            cLBSatz = frmWKL20.List3.Nodes(i).Text
        
        
            cArtNr = Mid(cLBSatz, 7, 6)
            byAWM = ermawm(Trim(cArtNr))
            
            
            If byAWM < 10 And byAWM > 0 Then
                frmWKL20.List3.Nodes(i).BackColor = glfarbe(byAWM)
                frmWKL20.List3.Nodes(i).ForeColor = vbBlack
            End If
            
            If byAWM < 20 And byAWM > 10 Then
                frmWKL20.List3.Nodes(i).BackColor = glfarbe2(byAWM - 10)
                frmWKL20.List3.Nodes(i).ForeColor = vbBlack
            End If
        
            If byAWM = 98 Then
                frmWKL20.List3.Nodes(i).ForeColor = vbRed
            End If
            
            If byAWM = 95 Then
                frmWKL20.List3.Nodes(i).BackColor = vbBlue
                frmWKL20.List3.Nodes(i).ForeColor = vbBlack
            End If
            
            If byAWM = 94 Then
                frmWKL20.List3.Nodes(i).BackColor = vbWhite
                frmWKL20.List3.Nodes(i).ForeColor = vbBlue
            End If
            
            If byAWM = 93 Then
                frmWKL20.List3.Nodes(i).BackColor = vbWhite
                frmWKL20.List3.Nodes(i).ForeColor = vbGreen
            End If
            
            If byAWM = 92 Then
                frmWKL20.List3.Nodes(i).BackColor = vbBlack
                frmWKL20.List3.Nodes(i).ForeColor = vbWhite
            End If
            
            If BistDuEinSonderkontingent(cArtNr) Then
                frmWKL20.List3.Nodes(i).BackColor = glfarbe(0)
            End If
            
        Next i
            
    ElseIf gsFARBKASSE = "2" Then
    
        Dim lAnzSatz    As Long
        Dim lAktSatz    As Long
    
        lAnzSatz = frmWKL20.List1.ListCount
    
        For lAktSatz = 0 To lAnzSatz - 1
        
            cLBSatz = frmWKL20.List1.list(lAktSatz)
        
            Dim lFarbcode As Long
            Dim lFarbe As Long
            Dim cBedienernummer As String
            Dim cSQL As String
            Dim rsrs As DAO.Recordset
            
            cBedienernummer = Mid(cLBSatz, 148, 3)
            cBedienernummer = Trim$(cBedienernummer)
            
            i = lAktSatz + 1

            cSQL = "Select * from BEDTERM where bednu =  " & cBedienernummer
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!FARBCODE) Then
                    lFarbcode = rsrs!FARBCODE
                Else
                    lFarbcode = 0
                End If
                Select Case lFarbcode
                    Case Is = 0
                        lFarbe = &H404040    'vbBlack
                        frmWKL20.List3.Nodes(i).ForeColor = vbWhite
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                    Case Is = 1
                        lFarbe = vbRed
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                    Case Is = 2
                        lFarbe = vbGreen
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 3
                        lFarbe = vbYellow
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 4
                        lFarbe = vbBlue
                        frmWKL20.List3.Nodes(i).ForeColor = vbWhite
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 5
                        lFarbe = vbMagenta
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 6
                        lFarbe = vbCyan
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 7
                        lFarbe = vbWhite
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 8
                        lFarbe = &HC0C0FF
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 9
                        lFarbe = &H40C0&
                        frmWKL20.List3.Nodes(i).ForeColor = vbWhite
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 10
                        lFarbe = &H80C0FF 'Apricot
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                    Case Is = 11
                        lFarbe = &HFF8080 '&H80000003 'Hellblau
                        frmWKL20.List3.Nodes(i).ForeColor = vbBlack
                        frmWKL20.List3.Nodes(i).BackColor = lFarbe
                        
                End Select
            End If
            rsrs.Close: Set rsrs = Nothing
    
        Next lAktSatz
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "farbelist3"
    Fehler.gsFehlertext = "Im Programmteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub farbelist4(frmx As Form)
    On Error GoTo LOKAL_ERROR
    
    Dim byAWM As Integer
    Dim i As Integer
    Dim j As Integer
    Dim cLBSatz As String
    Dim cArtNr As String
    Dim iStep As Integer
    
    Screen.MousePointer = 11
    
    iStep = 0
    
    For i = 1 To frmx.List4.Nodes.Count
        cLBSatz = frmx.List4.Nodes(i).Text
        
        iStep = 1
        
        cArtNr = Left(cLBSatz, 6)
        
        iStep = 2

        byAWM = Val(Mid(cLBSatz, 72, 2))
        
        iStep = 3
        
        If byAWM < 10 And byAWM > 0 Then
            frmx.List4.Nodes(i).BackColor = glfarbe(byAWM)
            frmx.List4.Nodes(i).ForeColor = vbBlack
        End If
        
        iStep = 4
        
        If byAWM < 20 And byAWM > 10 Then
            frmx.List4.Nodes(i).BackColor = glfarbe2(byAWM - 10)
            frmx.List4.Nodes(i).ForeColor = vbBlack
        End If
    
        iStep = 5
        
        If byAWM = 98 Then
            frmx.List4.Nodes(i).ForeColor = vbRed
        End If
        
        If byAWM = 95 Then
            frmx.List4.Nodes(i).BackColor = vbBlue
            frmx.List4.Nodes(i).ForeColor = vbWhite 'vbBlack
    
        End If
        
        If byAWM = 94 Then
            frmx.List4.Nodes(i).BackColor = vbWhite
            frmx.List4.Nodes(i).ForeColor = vbBlue
    
        End If
        
        If byAWM = 93 Then
            frmx.List4.Nodes(i).BackColor = vbWhite
            frmx.List4.Nodes(i).ForeColor = vbGreen
    
        End If
        
        If byAWM = 92 Then
            frmx.List4.Nodes(i).BackColor = vbBlack
            frmx.List4.Nodes(i).ForeColor = vbWhite
    
        End If
        
        iStep = 6
        
        If gbKONTIN Then
            If BistDuEinSonderkontingent(cArtNr) Then
                frmx.List4.Nodes(i).BackColor = vbGrayText ' glfarbe(0)
            End If
        End If
        
        iStep = 7
        
    Next i
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 13 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul20"
        Fehler.gsFunktion = "farbelist4"
        Fehler.gsFehlertext = "Im Programmteil Kasse/M20 ist ein Fehler aufgetreten." & iStep & " " & cArtNr & " " & byAWM
        
        Fehlermeldung1
    End If
End Sub
Public Function ermawm(cART As String) As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    ermawm = 0
    
    sSQL = "Select Awm from artikel where artnr = " & cART
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs!AWM) Then
            ermawm = Val(rs!AWM)
        End If
    End If

    rs.Close: Set rs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermawm"
    Fehler.gsFehlertext = "Im Programmteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub HangUp(ByVal Verbindung$)
    On Error GoTo LOKAL_ERROR
    
    Dim S As Long
      Dim l As Long
      Dim LN As Long
      Dim aa$
      ReDim R(255) As RASType 'RASCONN95
    
        R(0).dwSize = 412
        S = 256 * R(0).dwSize
        l = RasEnumConnections(R(0), S, LN)
        For l = 0 To LN - 1
          aa = StrConv(R(l).szEntryName(), vbUnicode)
          aa = Left$(aa, InStr(aa, Chr$(0)) - 1)
          If aa = Verbindung Then
            RCon = R(l).hRasCon
            Dim rec As Long
            rec = RasHangUp(RCon)
          End If
        Next l

'  Dim S As Long
'  Dim l As Long
'  Dim LN As Long, aa$
'  ReDim R(255) As RASType
'
'    R(0).dwSize = 412
'    S = 256 * R(0).dwSize
'    l = RasEnumConnections(R(0), S, LN)
'    For l = 0 To LN - 1
'      aa = StrConv(R(l).szEntryName(), vbUnicode)
'      aa = Left$(aa, InStr(aa, Chr$(0)) - 1)
'      If aa = Verbindung Then
'
'        RCon = R(l).hRasConn
'
'        Dim rec As Long
'        rec = RasHangUp(RCon)
'      End If
'    Next l
   
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "HangUp"
    Fehler.gsFehlertext = "Beim Beenden der DFÜ/Internet - Verbindung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Registrieren/DeRegistrieren einer ActiveX Komponente
'ohne REGSVR32.EXE, ohne Setup (VB5STKIT.DLL)
'
'(w) 15.01.2000 by Marcus Warm <mwarm@geosoft.de>
'
'Dateiname: Dateiname (.ocx, .dll)
'reg: True=Register, False=UnRegister
'
'Rückgabewerte:
'1  = Erfolg
'-2 = Datei nicht vorhanden oder keine gültige DLL
'     (LoadLibrary gescheitert)
'-3 = Adresse der DLL Funktion kann nicht ermittelt werden
'     (GetProcAddress gescheitert)
'-4 = Aufruf der DLL Funktion gescheitert
'     (CreateThread gescheitert)
'-5 = DLL Funktion gescheitert (möglicherweise Time-Out)
'     (WaitForSingleObject gescheitert)
'
Public Function Register(ByVal DateiName As String, ByVal reg As Boolean)
  Dim ret&, Library&, Func$, EntryPoint&, Thread&, R&

  On Error GoTo Fehler
    
  'Jede DLL und jedes OCX haben eine öffentliche
  'Library Funktion zum Registrieren der Komponente(n).
  'Eine ActiveX Komponente kann sich also selber
  'registrieren.
    
  'DLL laden ("Late Binding von Hand")
  Library = LoadLibrary(DateiName)
  ret = -2
  If Library Then
    If reg Then
      'Name der Funktion zum Registrieren
      Func = "DllRegisterServer"
    Else
      'Name der Funktion zum Deregistrieren
      Func = "DllUnregisterServer"
    End If
    'Adresse der DLL Funktion ermitteln (Late Binding!)
    EntryPoint = GetProcAddress(Library, Func)
    ret = -3
    If EntryPoint Then
      'DLL Funktion aufrufen
      Thread = CreateThread(ByVal 0, 0, ByVal EntryPoint, ByVal 0, 0, R)
      ret = -4
      If Thread Then
        'auf das Ende der DLL Funktion warten
        R = WaitForSingleObject(Thread, 10000)
        If R Then
          'Fehler!
          FreeLibrary Library
          'Thread abbrechen (Ressourcenfreigabe)
          'Handle für ExitThread ermitteln, damit der richtige
          'Thread abgebrochen wird:
          R = GetExitCodeThread(Thread, R)
          ExitThread R
          Register = -5
          Exit Function
        End If
        CloseHandle Thread
        ret = 1   'Erfolg!
      End If
    End If
    FreeLibrary Library
  End If
  
  Register = ret
Exit Function

Fehler:
  Register = err
End Function

Public Function FileExist(ByVal dn$) As Boolean

  On Error Resume Next
  FileExist = (Dir(dn) <> "")
  
End Function
Public Sub GeheAufStartModul20()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim iFehler As Integer
    
    For lcount = 0 To 17
        iFehler = lcount + 1
        frmWKL20!Command1(lcount).Visible = False
    Next lcount
    
iFehler = 19
    frmWKL20!Command1(11).Visible = True
iFehler = 20
    
    frmWKL20!Frame1.Visible = False
iFehler = 21
    
    frmWKL20!Label1(0).Visible = True
iFehler = 22
    frmWKL20!Text1(0).Visible = True
iFehler = 23
    frmWKL20!Label1(4).Visible = False
iFehler = 24
    frmWKL20!Label2(5).Visible = False
iFehler = 25
    frmWKL20!Label1(5).Visible = False
iFehler = 26
    frmWKL20!Label2(6).Visible = False
iFehler = 27
    frmWKL20!Label1(6).Visible = False
iFehler = 28
    frmWKL20!Text1(1).Visible = False
iFehler = 29
    frmWKL20!Label8(2).Caption = ""
iFehler = 30
    frmWKL20!Label8(3).Caption = ""
iFehler = 31
    
    frmWKL20!Label1(8).Caption = ""
iFehler = 32
    
    frmWKL20!Text1(0).Text = ""
iFehler = 33

iFehler = 34
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "GeheAufStartModul20"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten." & Trim$(Str$(iFehler))
    
    Fehlermeldung1
End Sub
Public Sub LeereDialogModul20()
    On Error GoTo LOKAL_ERROR
    
    'Sonderpreis
    frmWKL20!Label1(1).Visible = False
    frmWKL20!Label2(0).Caption = "0,00"
    frmWKL20!Label2(0).Visible = False
    
    'nettospannenanzeige
    frmWKL20!Shape1.Visible = False
    frmWKL20!Shape2.Visible = False
    frmWKL20!Shape3.Visible = False
            
    frmWKL20!Label39(3).Caption = ""
    frmWKL20!Label39(4).Caption = ""
    frmWKL20!Label39(0).Caption = ""

    'Artikelrabatt
    If gbArtrabhalten = True Then
    
    Else
        frmWKL20!Label1(2).Visible = False
        frmWKL20!Label2(1).Caption = "0,00"
        frmWKL20!Label2(1).Visible = False
        frmWKL20!Label2(2).Visible = False
    End If
    
    'PreisKennzeichen
    giPreisKz = 0
    
    'Menge
    frmWKL20!Label2(5).Caption = "1"
    
    'Kundennummer
    frmWKL20.kundenauswahlausblend
    
    'Globale Variablen leeren
    gcRueckgeld = ""
    gdGegeben = 0
    gdSumme = 0
    
    'auf normalen Preis setzen
    frmWKL20!Label8(3).Caption = 0
    
    'Rückgutschein auf leer
    gcRueckGutsch = ""
    
    'Display auf nächsten Kunden
'    InitKundenDisplayModul20
    
    frmWKL20!Text1(1).Text = ""
    frmWKL20.Caption = ""
    
    frmWKL20!Picture2.Tag = ""
    frmWKL20!Picture2.Visible = False
    
    DoEvents
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LeereDialogModul20"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermartnrausWGN(cWg As String) As String
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermartnrausWGN = ""

If IsNumeric(cWg) Then
    sSQL = " Select artnr from warengru where wgnr = " & cWg
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!artnr) Then
            ermartnrausWGN = rsrs!artnr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermartnrausWGN"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function IstDasEineWGN(lartnr As Long) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    IstDasEineWGN = False

    sSQL = " Select * from warengru where artnr = " & lartnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        IstDasEineWGN = True
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "IstDasEineWGN"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function IstDasEinAktiverTerminArtikel(lartnr As Long) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    IstDasEinAktiverTerminArtikel = False

    sSQL = " Select * from PRSTERM where artnr = " & lartnr & " and status = 1 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        IstDasEinAktiverTerminArtikel = True
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "IstDasEinAktiverTerminArtikel"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function IstDasEineWGNforKasse(cArtNr As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    IstDasEineWGNforKasse = False
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    If Len(cArtNr) > 6 Then
        Exit Function
    End If

    sSQL = " Select * from warengru where artnr = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        IstDasEineWGNforKasse = True
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "IstDasEineWGNforKasse"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function Welche_Warengruppen_Taste(lartnr As Long) As Integer
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    Welche_Warengruppen_Taste = -1

    sSQL = " Select wgnr from warengru where artnr = " & lartnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!WGNR) Then
            If rsrs!WGNR <= 12 Then
                Welche_Warengruppen_Taste = rsrs!WGNR + 17
            Else
                Welche_Warengruppen_Taste = rsrs!WGNR + 107
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Welche_Warengruppen_Taste"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function IstArtikelnichtStornierfähig(cArtNr As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    IstArtikelnichtStornierfähig = False

    sSQL = " Select * from Stornof where artnr = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        IstArtikelnichtStornierfähig = True
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "IstArtikelnichtStornierfähig"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub InitKundenDisplayModul20()
    On Error GoTo LOKAL_ERROR
    
    Dim czeit As String
    Dim cHH As String
    Dim cZeile1 As String
    Dim cZeile2 As String
    
    If gbDisplay Or gbZweitMoni Then
        cZeile1 = Chr$(31) & Chr$(67) + "0"
        cZeile2 = ""
        ZeigeKundenDisplay cZeile1, cZeile2, "", "", -1
    Else
        Exit Sub
    End If
    
    If frmWKL20!Label2(6).Caption = "0,00" Then
        czeit = Format$(Now, "HH:MM")
        cHH = Left(czeit, 2)
        If Val(cHH) < 11 Then
            cZeile1 = gsMORGENTEXT '"Guten Morgen!"
        ElseIf Val(cHH) < 18 Then
            cZeile1 = gsMITTAGTEXT '"Guten Tag!          "
        Else
            cZeile1 = gsABENDTEXT '"Guten Abend!        "
        End If
        cZeile2 = Space$(20)
    Else
        cZeile1 = Space$(20)
        cZeile2 = Space$(20)
    End If
    
    
    If gbDisplay Or gbZweitMoni Then
    
        If gbZweitMoniMinimieren = False Then
            ZeigeKundenDisplay cZeile1, cZeile2
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "InitKundenDisplayModul20"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub InsertAFCBuchGutscheinModul20(dZhlgGutsch As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim dLiNr As Double
    Dim dEkpr As Double
    Dim dWert As Double
    Dim ctmp As String
    Dim cLBSatz As String
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsKJ As Recordset
    Dim rsAFCGR As Recordset
    Dim iFeld As Integer
    Dim iDbNr As Integer
    
    '*** Kunden-Umsatz ****
    Dim dKdUmsatz As Double
    Dim dKdBonus As Double

    '*** KASSJOUR-Felder ****
    Dim cKJArtNr As String
    Dim cKJBezeich As String
    Dim cKJMenge As String
    Dim dKJPreis As Double
    Dim lKJADate As Long
    Dim cKJAZeit As String
    Dim lKJBediener As Long
    Dim cKJKundNr As String
    Dim cKJFiliale As String
    Dim cKJKasNum As String
    Dim cKJLiNr As String
    Dim cKJLPZ As String
    Dim cKJAGN As String
    Dim cKJEAN As String
    Dim cKJMwst As String
    Dim dKJEkpr As Double
    Dim dKJVkpr As Double
    Dim cKJBelegNr As String
    Dim dKJBest1 As Double
    Dim sKK_art As String
    Dim cUmsOK As String
    Dim sPreisKz As String
    
    Dim cWasSuchteMan As String
    Dim lPos As Long
    
    Dim cpfaddb As String
        
    cpfaddb = gcDBPfad
    If Right$(cpfaddb, 1) <> "\" Then
        cpfaddb = cpfaddb & "\"
    End If
    
    iFeld = 111
    
    ctmp = Trim$(frmWKL20!Label2(7).Caption)
    If Val(ctmp) < 0 Then
        ctmp = "0"
    End If
    cKJKundNr = ctmp
    
    If Val(cKJKundNr) > 0 Then
        'dann nach Preiskz fragen
        sPreisKz = ermPREISKZ(cKJKundNr)
    End If
    
    
    
    
    
    cSQL = "Select * from KASSJOUR where ARTNR = -1"
    Set rsKJ = gdBase.OpenRecordset(cSQL)
    iFeld = 112
    lAnzSatz = frmWKL20!List1.ListCount
    iFeld = 113
    For lAktSatz = 0 To lAnzSatz - 1
    
        iFeld = 1
        cKJArtNr = ""
        cKJBezeich = ""
        cKJMenge = ""
        dKJPreis = 0
        lKJADate = 0
        cKJAZeit = ""
        lKJBediener = 0
        cKJKundNr = ""
        cKJFiliale = ""
        cKJKasNum = ""
        cKJLiNr = ""
        cKJLPZ = ""
        cKJAGN = ""
        cKJEAN = ""
        cKJMwst = ""
        dKJEkpr = 0
        dKJVkpr = 0
        cKJBelegNr = ""
        dKJBest1 = 0
        
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        '//2002
        If Len(cLBSatz) > 156 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If
        
        
        
        'Grund.cfg Teil1
        
        If FileExists(cpfaddb & "Grund.cfg") Then
        
            lPos = 0
            lPos = InStr(cLBSatz, "@Q")
        
            If lPos > 0 Then
                cWasSuchteMan = Trim(Mid(cLBSatz, lPos + 2, Len(cLBSatz) - lPos))
            End If
            
        End If
        'Ende Grund.cfg
        
        
        

        
        ctmp = Mid(cLBSatz, 7, 6)
        ctmp = Trim$(ctmp)
        
        'Zeile ZWISCHENSUMME darf nicht übernommen werden!
        If ctmp <> "000000" Then
            cSQL = "Select * from ARTIKEL where ARTNR = " & ctmp
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                
                If Not IsNull(rsrs!UMS_OK) Then
                    cUmsOK = rsrs!UMS_OK
                Else
                    cUmsOK = "J"
                End If
                
                iFeld = 2
                If Not IsNull(rsrs!LPZ) Then
                    cKJLPZ = rsrs!LPZ
                Else
                    cKJLPZ = ""
                End If
                
                iFeld = 3
                If Not IsNull(rsrs!AGN) Then
                    cKJAGN = rsrs!AGN
                Else
                    cKJAGN = ""
                End If
                
                iFeld = 4
                If Not IsNull(rsrs!EAN) Then
                    cKJEAN = rsrs!EAN
                Else
                    cKJEAN = ""
                End If
                
                iFeld = 5
                If Not IsNull(rsrs!lekpr) Then
                    dEkpr = rsrs!lekpr
                Else
                    dEkpr = 0
                End If
                
                iFeld = 6
                If Not IsNull(rsrs!linr) Then
                    dLiNr = rsrs!linr
                Else
                    dLiNr = 0
                End If
            Else
                dEkpr = 0
                dLiNr = 0
            End If
            rsrs.Close: Set rsrs = Nothing
            
            cSQL = "Select * from AFCBUCH where AARTNR = -1"
            Set rsrs = gdBase.OpenRecordset(cSQL)
            
            rsrs.AddNew
            rsrs!SYNStatus = "A"
            iFeld = 7
            ctmp = Mid(cLBSatz, 148, 3)
            ctmp = Trim$(ctmp)
            rsrs!abednu = Val(ctmp)
            lKJBediener = Val(ctmp)
            
            iFeld = 8
            rsrs!AFLAG = 0
            
            iFeld = 9
            
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If
            
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!aMenge = Val(ctmp)
            cKJMenge = ctmp
            
            
            iFeld = 10
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!APREIS = Val(ctmp)
            dKJPreis = rsrs!APREIS
            
            iFeld = 11
            ctmp = Mid(cLBSatz, 7, 6)
            ctmp = Trim$(ctmp)
            rsrs!aartnr = Val(ctmp)
            cKJArtNr = ctmp
            
            iFeld = 12
            ctmp = Mid(cLBSatz, 14, 35)
            ctmp = Trim$(ctmp)
            rsrs!ABEZEICH = ctmp
            cKJBezeich = ctmp
            
            iFeld = 13
            rsrs!ADATE = Fix(Now)
            
            iFeld = 14
            rsrs!AZEIT = Format$(Now, "HH:MM:SS")
            lKJADate = rsrs!ADATE
            cKJAZeit = rsrs!AZEIT
            
            iFeld = 15
            ctmp = Mid(cLBSatz, 72, 1)
            ctmp = Trim$(ctmp)
            rsrs!AMWSK = ctmp
            cKJMwst = ctmp
            
            iFeld = 16
            If ctmp = "V" Then
                ctmp = Mid(cLBSatz, 104, 9)
            ElseIf ctmp = "E" Then
                ctmp = Mid(cLBSatz, 114, 9)
            Else
                ctmp = "0"
            End If
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!AMWST = Val(ctmp)
            
            iFeld = 17
            ctmp = frmWKL20!Label2(7).Caption
            ctmp = Trim$(ctmp)
            If Val(ctmp) < 0 Then
                ctmp = "0"
            End If
            rsrs!AKUNUM = Val(ctmp)
            cKJKundNr = ctmp
            
            iFeld = 18
            rsrs!BELEGNR = gdBonNr
            cKJBelegNr = rsrs!BELEGNR
            
            iFeld = 19
            'rsRs!KASNUM = 1
            rsrs!kasnum = gcKasNum
            'cKJKasNum = "1"
            cKJKasNum = gcKasNum
            cKJFiliale = gcFilNr
            
            iFeld = 20
            rsrs!BUCHFLAG = 0
            
            iFeld = 21
            ctmp = Mid(cLBSatz, 50, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!AALTPREIS = Val(ctmp)
            
            iFeld = 22
            ctmp = Mid(cLBSatz, 128, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            '//Aenderung
            If ctmp = 0 Then
                ctmp = rsrs!APREIS / rsrs!aMenge
                ctmp = Trim$(ctmp)
                ctmp = fnMoveComma2Point$(ctmp)
            End If
            rsrs!AVKPR = Val(ctmp)
            dKJVkpr = rsrs!AVKPR
            
            iFeld = 23
            rsrs!ALEKPR = dEkpr
            dKJEkpr = dEkpr
            
            iFeld = 24
            rsrs!linr = dLiNr
            cKJLiNr = Trim$(Str$(dLiNr))
            
            iFeld = 25
            rsrs!kk_art = "MX"
            
            iFeld = 26
            ctmp = Mid(cLBSatz, 138, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dKJBest1 = Val(ctmp)
            rsrs!BESTAND = Val(ctmp)
            
            rsrs!ZHLGGUTSCH = dZhlgGutsch
            
            
'            'ist Preiskz = 6 also Netto dann ums_ok = N
'            If Val(sPreisKz) = 6 Then
'                cUmsOK = "N"
'            End If
            
            rsrs!UMS_OK = cUmsOK
            rsrs!FILIALNR = Val(gcFilNr)
            rsrs.Update   '** updating AFCBUCH **
            
            rsrs.Close: Set rsrs = Nothing
            
            
            
            
            
            'Grund.cfg Teil2
            If FileExists(cpfaddb & "Grund.cfg") Then
            
                
            
                Set rsAFCGR = gdBase.OpenRecordset("AFCBUCH_GRUND", dbOpenTable)
    
                rsAFCGR.AddNew
                rsAFCGR!BEZEICH = cKJBezeich
                rsAFCGR!Menge = Val(cKJMenge)
                rsAFCGR!ADATE = lKJADate
                rsAFCGR!AZEIT = cKJAZeit
                
                If IsNumeric(cWasSuchteMan) = True Then
                    rsAFCGR!EAN = cWasSuchteMan
                Else
                    rsAFCGR!EAN = cKJArtNr
                End If
                    
                rsAFCGR.Update
                rsAFCGR.Close: Set rsAFCGR = Nothing
                

            End If
            'Ende Grund.cfg
            
            'doemer.cfg
            If FileExists(cpfaddb & "doemer.cfg") Then
                doemer_bestand_updaten cKJArtNr, Val(cKJMenge)
            End If
            'Ende doemer.cfg
    
            
            
            
            
            
            '** Tabelle KASSJOUR **
            rsKJ.AddNew
            iFeld = 26
            rsKJ!artnr = Val(cKJArtNr)
            iFeld = 27
            rsKJ!BEZEICH = cKJBezeich
            iFeld = 28
            rsKJ!Menge = Val(cKJMenge)
            iFeld = 29
            rsKJ!Preis = dKJPreis
            iFeld = 30
            rsKJ!ADATE = lKJADate
            iFeld = 31
            rsKJ!AZEIT = cKJAZeit
            iFeld = 32
            rsKJ!BEDIENER = lKJBediener
            iFeld = 33
            rsKJ!Kundnr = Val(cKJKundNr)
            iFeld = 34
            rsKJ!FILIALE = Val(cKJFiliale)
            iFeld = 35
            rsKJ!kasnum = Val(cKJKasNum)
            iFeld = 36
            rsKJ!linr = Val(cKJLiNr)
            iFeld = 37
            rsKJ!LPZ = Val(cKJLPZ)
            iFeld = 38
            rsKJ!AGN = Val(cKJAGN)
            iFeld = 39
            rsKJ!EAN = cKJEAN
            iFeld = 40
            rsKJ!MWST = cKJMwst
            iFeld = 41
            rsKJ!ekpr = dKJEkpr
            iFeld = 42
            rsKJ!vkpr = dKJVkpr
            iFeld = 43
            rsKJ!BELEGNR = gdBonNr
            iFeld = 44
            
'            'ist Preiskz = 6 also Netto dann ums_ok = N
'            If Val(sPreisKz) = 6 Then
'                cUmsOK = "N"
'            End If
            
            rsKJ!UMS_OK = cUmsOK
            iFeld = 45
            sKK_art = "MX"
            rsKJ!kk_art = sKK_art
            rsKJ!best1 = dKJBest1
            rsKJ.Update
            
            If Val(cKJArtNr) <> 666666 Then
                If cUmsOK <> "N" Then
                    dKdUmsatz = dKdUmsatz + dKJPreis
                Else
                
                End If
            End If
            iFeld = 46
            If Len(cLBSatz) >= 154 Then
                If Mid(cLBSatz, 154, 1) <> "N" Then
                    dKdBonus = dKdBonus + dKJPreis
                End If
            Else
                dKdBonus = dKdBonus + dKJPreis
            End If
        End If
        
    Next lAktSatz
    rsKJ.Close
    
    iFeld = 47
    
    If Val(cKJKundNr) > 0 Then
        Plus_Bonus cKJKundNr, dKdUmsatz, dKdBonus
    End If
    
    iFeld = 48
    If gbbonusausjetzt = True Then
    
        Dim Bonusneu        As String
        Dim Bonusalt        As String
        Dim cDatum          As String
        Dim czeit           As String
        
        Dim cVname          As String
        Dim cNName           As String
        
        ReDim cZeilen(0 To 9) As String
        iFeld = 49
        cSQL = "Select * from Kunden where kundnr = " & cKJKundNr
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.Edit
            rsrs!Status = "E"
            rsrs!SYNStatus = "E"
            
            If Not IsNull(rsrs!BONUS) Then
                Bonusalt = Format(rsrs!BONUS, "##000.00 EUR")
                rsrs!BONUS = rsrs!BONUS - gdbonusHerabwert
                Bonusneu = Format(rsrs!BONUS, "##000.00 EUR")
            Else
                Bonusalt = "000.00 EUR"
                rsrs!BONUS = gdbonusHerabwert * (-1)
                Bonusneu = Format(rsrs!BONUS, "##000.00 EUR")
            End If
            
            If Not IsNull(rsrs!TBONUS) Then
                rsrs!TBONUS = CDbl(rsrs!TBONUS) - gdbonusHerabwert
            Else
                rsrs!TBONUS = gdbonusHerabwert * (-1)
            End If
            
            rsrs!LASTDATE = DateValue(Now)
            rsrs!LASTTIME = TimeValue(Now)
            rsrs.Update
            
            iFeld = 50
            schreibeKBProtokoll Space(10 - Len(cKJKundNr)) & cKJKundNr & " Bonus reduziert jetzt: " & Bonusneu & " vorher: " & Bonusalt
            
            cDatum = DateValue(Now)
            czeit = TimeValue(Now)
            
            iFeld = 51
            cVname = lookingForKundendaten(cKJKundNr).vorname
            cNName = lookingForKundendaten(cKJKundNr).nachname
            
            cZeilen(0) = "Bonus reduziert"
            cZeilen(1) = "-----------------"
            cZeilen(2) = "KundNr:  " & cKJKundNr
            cZeilen(3) = "Vorname: " & cVname
            cZeilen(4) = "Name:    " & cNName
            cZeilen(5) = "vorher:  " & Bonusalt
            cZeilen(6) = "jetzt:   " & Bonusneu
            cZeilen(7) = "Datum:   " & cDatum
            cZeilen(8) = "Zeit:    " & czeit
            iFeld = 52
            DruckeArbeitszeitBelegWK20d cZeilen(), 8
            If gb2BONUSMESS Then
                DruckeArbeitszeitBelegWK20d cZeilen(), 8
            End If
        End If
        
        rsrs.Close: Set rsrs = Nothing

    End If
    
    Dim cGutsch As String
    Dim dGeldwert As Double
    Dim i As Integer
    If frmWK20g.List3.ListCount > 0 Then
        For i = 0 To frmWK20g.List3.ListCount - 1
            cGutsch = ""
            cGutsch = Val(Left(frmWK20g.List3.list(i), 8))
            dGeldwert = CDbl(Mid(frmWK20g.List3.list(i), 20, 11))
            insertGUTZ Fix(Now), Format$(Now, "HH:MM:SS"), CStr(gdBonNr), gcKasNum, "EI", dGeldwert, cGutsch, CStr(lKJBediener)
        Next i
    End If
    
    gdbonusHerabwert = 0
    gbbonusausjetzt = False
    gbbonusHerab = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "InsertAFCBuchGutscheinModul20"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten." & Trim$(Str$(iFeld))
    
    Fehlermeldung1
    
End Sub
Public Sub Plus_Bonus(cKundnr As String, dKundenumsatz As Double, dKundenbonus As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select UMSLJ,BONUS,TBONUS, LASTDATE, LASTTIME, SYNStatus, Status from KUNDEN where KUNDNR = " & cKundnr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
        If Not IsNull(rsrs!UMSLJ) Then
            rsrs!UMSLJ = rsrs!UMSLJ + dKundenumsatz
        Else
            rsrs!UMSLJ = dKundenumsatz
        End If
        
        If gbKUBONUS Then
        
            If Not IsNull(rsrs!BONUS) Then
                rsrs!BONUS = rsrs!BONUS + dKundenbonus
            Else
                rsrs!BONUS = dKundenbonus
            End If
            
            If Not IsNull(rsrs!TBONUS) Then
                rsrs!TBONUS = CDbl(rsrs!TBONUS) + dKundenbonus
            Else
                rsrs!TBONUS = dKundenbonus
            End If
        End If
        
        rsrs!LASTDATE = DateValue(Now)
        rsrs!LASTTIME = TimeValue(Now)
        rsrs!SYNStatus = "E"
        rsrs!Status = "E"
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Plus_Bonus"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'Public Sub Bonus_runter(cKundnr As String, dKundenbonus As Double)
'On Error GoTo LOKAL_ERROR
'
'    Dim cSQL As String
'    Dim rsrs As Recordset
'
'    cSQL = "Select * from KUNDEN where KUNDNR = " & cKundnr & " "
'    Set rsrs = gdBase.OpenRecordset(cSQL)
'    If Not rsrs.EOF Then
'        rsrs.Edit
'
'        If gbKUBONUS Then
'            If Not IsNull(rsrs!BONUS) Then
'                rsrs!BONUS = rsrs!BONUS + dKundenbonus
'            Else
'                rsrs!BONUS = dKundenbonus
'            End If
'
'            If Not IsNull(rsrs!TBONUS) Then
'                rsrs!TBONUS = CDbl(rsrs!TBONUS) + dKundenbonus
'            Else
'                rsrs!TBONUS = dKundenbonus
'            End If
'        End If
'
'        rsrs!LASTDATE = DateValue(Now)
'        rsrs!LASTTIME = TimeValue(Now)
'        rsrs!SYNStatus = "E"
'        rsrs!Status = "E"
'        rsrs.Update
'    End If
'    rsrs.Close: Set rsrs = Nothing
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = "Modul2"
'    Fehler.gsFunktion = "Bonus_runter"
'    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Public Sub BonusVeränderung(cART As String, lkunde As Long, ByRef dKdBonus As Double, ByRef dKdUmsatz As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dKUmsatz    As Double
    Dim dKBonus     As Double
    
    If cART = "positiv" Then
        dKUmsatz = dKdUmsatz
        dKBonus = dKdBonus
    Else
        dKUmsatz = dKdUmsatz * (-1)
        dKBonus = dKdBonus * (-1)
    End If
    
    If lkunde > 0 Then
        Plus_Bonus CStr(lkunde), dKUmsatz, dKBonus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Bonusveränderung"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub InsertAFCBuchModul20For68()
    On Error GoTo LOKAL_ERROR
    
    Dim dLiNr As Double
    Dim dEkpr As Double
    Dim dWert As Double
    Dim ctmp As String
    Dim cLBSatz As String
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsKJ As Recordset
    Dim cGZart As String
    Dim dGeldwert As Double
    Dim iFeld As Integer
    Dim iDbNr As Integer
    
    
    '*** Kunden-Umsatz ****
    Dim dKdUmsatz As Double
    Dim dKdBonus As Double

    '*** KASSJOUR-Felder ****
    Dim cKJArtNr As String
    Dim cKJBezeich As String
    Dim cKJMenge As String
    Dim dKJPreis As Double
    Dim lKJADate As Long
    Dim cKJAZeit As String
    Dim lKJBediener As Long
    Dim cKJKundNr As String
    Dim cKJFiliale As String
    Dim cKJKasNum As String
    Dim cKJLiNr As String
    Dim cKJLPZ As String
    Dim cKJAGN As String
    Dim cKJEAN As String
    Dim cKJMwst As String
    Dim dKJEkpr As Double
    Dim dKJVkpr As Double
    Dim cKJBelegNr As String
    Dim dKJBest1 As Double
    Dim sKK_art As String
    Dim cUmsOK As String
    Dim cKJMopreis As String
    Dim cExtend As String
    Dim cGutsch As String
    Dim cBonusOk As String
    Dim cArtMWSt    As String
    Dim dSpanne     As Double
    Dim sArtnr As String
    
    Dim cWasSuchteMan As String
    Dim lPos As Long
    Dim sPreisKz As String
    Dim cpfaddb As String
        
    cpfaddb = gcDBPfad
    If Right$(cpfaddb, 1) <> "\" Then
        cpfaddb = cpfaddb & "\"
    End If
    
    
    iFeld = 111
    
            
    ctmp = Trim$(frmWKL68.Label1(27).Caption)
    If Val(ctmp) < 0 Then
        ctmp = "0"
    End If
    cKJKundNr = ctmp
    
    If Val(cKJKundNr) > 0 Then
        'dann nach Preiskz fragen
        sPreisKz = ermPREISKZ(cKJKundNr)
    End If

    
    
    
    
    
   
    iFeld = 112
    lAnzSatz = frmWKL20!List1.ListCount
    iFeld = 113
    For lAktSatz = 0 To lAnzSatz - 1
    
        iFeld = 1
        cKJArtNr = ""
        cKJBezeich = ""
        cKJMenge = ""
        dKJPreis = 0
        lKJADate = 0
        cKJAZeit = ""
        lKJBediener = 0
        cKJKundNr = ""
        cKJFiliale = ""
        cKJKasNum = ""
        cKJLiNr = ""
        cKJLPZ = ""
        cKJAGN = ""
        cKJEAN = ""
        cKJMwst = ""
        dKJEkpr = 0
        dKJVkpr = 0
        cKJBelegNr = ""
        dKJBest1 = 0
        
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        'Besonderheiten am Satzende

        'hier Besonders Merkmal - wird in Mopreis kassjour gespeichert
        
        If Len(cLBSatz) > 175 Then
            cKJMopreis = Mid(cLBSatz, 177, 8)
        Else
            cKJMopreis = "0"
        End If

        If Len(cLBSatz) > 157 Then
            cExtend = Mid(cLBSatz, 158, 18)
        Else
            cExtend = ""
        End If
        
        'Grund.cfg Teil1
        
        If FileExists(cpfaddb & "Grund.cfg") Then
        
            lPos = 0
            lPos = InStr(cLBSatz, "@Q")
        
            If lPos > 0 Then
                cWasSuchteMan = Trim(Mid(cLBSatz, lPos + 2, Len(cLBSatz) - lPos))
            End If
            
        End If
        'Ende Grund.cfg
        
        


        
        ctmp = Mid(cLBSatz, 7, 6)
        ctmp = Trim$(ctmp)
        sArtnr = ctmp
        
        'Zeile ZWISCHENSUMME darf nicht übernommen werden!
        If ctmp <> "000000" Then
            cSQL = "Select * from ARTIKEL where ARTNR = " & ctmp
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                
                If Not IsNull(rsrs!UMS_OK) Then
                    cUmsOK = rsrs!UMS_OK
                Else
                    cUmsOK = "J"
                End If
                
                iFeld = 2
                If Not IsNull(rsrs!LPZ) Then
                    cKJLPZ = rsrs!LPZ
                Else
                    cKJLPZ = ""
                End If
                
                iFeld = 3
                If Not IsNull(rsrs!AGN) Then
                    cKJAGN = rsrs!AGN
                Else
                    cKJAGN = ""
                End If
                
                iFeld = 4
                If Not IsNull(rsrs!EAN) Then
                    cKJEAN = rsrs!EAN
                Else
                    cKJEAN = ""
                End If
                
                iFeld = 5
                If Not IsNull(rsrs!ekpr) Then
                    dEkpr = rsrs!ekpr
                Else
                    dEkpr = 0
                End If
                
                iFeld = 6
                If Not IsNull(rsrs!linr) Then
                    dLiNr = rsrs!linr
                Else
                    dLiNr = 0
                End If
                
                If Not IsNull(rsrs!MWST) Then
                    cArtMWSt = rsrs!MWST
                Else
                    cArtMWSt = "V"
                End If
                
                'ist Preiskz = 6 also Netto dann mwst = O
                If Val(sPreisKz) = 6 Then
                    cArtMWSt = "O"
                End If
                
                
                
                iFeld = 9
                If Not IsNull(rsrs!BONUS_OK) Then
                    cBonusOk = rsrs!BONUS_OK
                Else
                    cBonusOk = "J"
                End If
                
                
                If Not IsNull(rsrs!SPANNE) Then
                    dSpanne = rsrs!SPANNE
                Else
                    dSpanne = 0
                End If
                
            Else
                dEkpr = 0
                dLiNr = 0
            End If
            
            rsrs.Close: Set rsrs = Nothing
                    
            
            cSQL = "Select * from AFCBUCH where AARTNR = -1"
            FnOpenrecordset rsrs, cSQL, 1, gdBase

            rsrs.AddNew
            rsrs!SYNStatus = "A"
            
            If ctmp = "666666" Then
                If gbGutscheinBeiVKversteuern = True Then
                    cBonusOk = "N"
                    cUmsOK = "J"
                    cArtMWSt = "V"
                Else
                    cBonusOk = "N"
                    cUmsOK = "N"
                    cArtMWSt = "O"
                End If
            End If
            
            iFeld = 7
            ctmp = Mid(cLBSatz, 148, 3)
            ctmp = Trim$(ctmp)
            
           
            rsrs!abednu = Val(ctmp)
            lKJBediener = Val(ctmp)
            
            iFeld = 8
            rsrs!AFLAG = 0
            
            iFeld = 9
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!aMenge = Val(ctmp)
            cKJMenge = ctmp
            
            
            iFeld = 10
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!APREIS = Val(ctmp)
            dKJPreis = rsrs!APREIS
            
            iFeld = 11
            ctmp = Mid(cLBSatz, 7, 6)
            ctmp = Trim$(ctmp)
            rsrs!aartnr = Val(ctmp)
            cKJArtNr = ctmp
            
            iFeld = 12
            ctmp = Mid(cLBSatz, 14, 35)
            ctmp = Trim$(ctmp)
            rsrs!ABEZEICH = ctmp
            cKJBezeich = ctmp
            
            iFeld = 13
            rsrs!ADATE = Fix(Now)
            
            iFeld = 14
            rsrs!AZEIT = Format$(Now, "HH:MM:SS")
            lKJADate = rsrs!ADATE
            cKJAZeit = rsrs!AZEIT
            

            
            rsrs!AMWSK = cArtMWSt
            cKJMwst = cArtMWSt
            
            
            iFeld = 16
            If ctmp = "V" Then
                ctmp = Mid(cLBSatz, 104, 9)
            ElseIf ctmp = "E" Then
                ctmp = Mid(cLBSatz, 114, 9)
            Else
                ctmp = "0"
            End If
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!AMWST = Val(ctmp)
            
            iFeld = 17
            ctmp = frmWKL68.Label1(27).Caption
            ctmp = Trim$(ctmp)
            If Val(ctmp) < 0 Then
                ctmp = "0"
            End If
            rsrs!AKUNUM = Val(ctmp)
            cKJKundNr = ctmp
            
'            'ist Preiskz = 6 also Netto dann ums_ok = N
'            If Val(sPreisKz) = 6 Then
'                cUmsOK = "N"
'            End If
            
            
            

            
            
            
            
            iFeld = 18
            rsrs!BELEGNR = gdBonNr
            cKJBelegNr = rsrs!BELEGNR
            
            iFeld = 19
           
            rsrs!kasnum = gcKasNum
            cKJKasNum = gcKasNum
            cKJFiliale = gcFilNr
            
            iFeld = 20
            rsrs!BUCHFLAG = 0
            
            iFeld = 21
            ctmp = Mid(cLBSatz, 50, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!AALTPREIS = Val(ctmp)
            
            iFeld = 22
            ctmp = Mid(cLBSatz, 128, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            '//Aenderung
            If ctmp = 0 Then
                ctmp = rsrs!APREIS / rsrs!aMenge
                ctmp = Trim$(ctmp)
                ctmp = fnMoveComma2Point$(ctmp)
            End If
            rsrs!AVKPR = Val(ctmp)
            dKJVkpr = rsrs!AVKPR
            
            iFeld = 23
            
        
            If dEkpr = 0 Then
            
                If sArtnr = "666668" Or sArtnr = "666669" Then
                    If gdZeitungsSpanne <> 0 Then
                        dEkpr = EKausNettospanneerrechnen(gdZeitungsSpanne, Val(ctmp), cArtMWSt)
                    End If
                Else
                    If dSpanne <> 0 Then
                        dEkpr = EKausNettospanneerrechnen(dSpanne, Val(ctmp), cArtMWSt)
                    End If
                End If
            
            End If
            
            rsrs!ALEKPR = dEkpr
            dKJEkpr = dEkpr
            
            iFeld = 24
            rsrs!linr = dLiNr
            cKJLiNr = Trim$(Str$(dLiNr))
            
            iFeld = 25
            rsrs!kk_art = "GZ"
            
            iFeld = 26
            ctmp = Mid(cLBSatz, 138, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dKJBest1 = Val(ctmp)
            rsrs!BESTAND = Val(ctmp)
            
            rsrs!ZHLGGUTSCH = 0 ' dZhlgGutsch
            rsrs!UMS_OK = cUmsOK
            rsrs!FILIALNR = Val(gcFilNr)
            rsrs!BONUS_OK = cBonusOk
            rsrs.Update   '** updating AFCBUCH **
            
            rsrs.Close: Set rsrs = Nothing
            
            
            
            'Grund.cfg Teil2
            If FileExists(cpfaddb & "Grund.cfg") Then
            
                Set rsKJ = gdBase.OpenRecordset("AFCBUCH_GRUND", dbOpenTable)
    
                rsKJ.AddNew
                rsKJ!BEZEICH = cKJBezeich
                rsKJ!Menge = Val(cKJMenge)
                rsKJ!ADATE = lKJADate
                rsKJ!AZEIT = cKJAZeit
                If IsNumeric(cWasSuchteMan) = True Then
                    rsKJ!EAN = cWasSuchteMan
                Else
                    rsKJ!EAN = cKJArtNr
                End If
                rsKJ.Update
                rsKJ.Close: Set rsKJ = Nothing
            
            End If
            'Ende Grund.cfg
            
            'doemer.cfg
            If FileExists(cpfaddb & "doemer.cfg") Then
                doemer_bestand_updaten cKJArtNr, Val(cKJMenge)
            End If
            'Ende doemer.cfg
    
            cSQL = "Select * from KASSJOUR where ARTNR = -1"
            FnOpenrecordset rsKJ, cSQL, 1, gdBase
            
            '** Tabelle KASSJOUR **
            rsKJ.AddNew
            iFeld = 26
            rsKJ!artnr = Val(cKJArtNr)
            iFeld = 27
            rsKJ!BEZEICH = cKJBezeich
            iFeld = 28
            rsKJ!Menge = Val(cKJMenge)
            iFeld = 29
            rsKJ!Preis = dKJPreis
            iFeld = 30
            rsKJ!ADATE = lKJADate
            iFeld = 31
            rsKJ!AZEIT = cKJAZeit
            iFeld = 32
            rsKJ!BEDIENER = lKJBediener
            iFeld = 33
            rsKJ!Kundnr = Val(cKJKundNr)
            iFeld = 34
            rsKJ!FILIALE = Val(cKJFiliale)
            iFeld = 35
            rsKJ!kasnum = Val(cKJKasNum)
            iFeld = 36
            rsKJ!linr = Val(cKJLiNr)
            iFeld = 37
            rsKJ!LPZ = Val(cKJLPZ)
            iFeld = 38
            rsKJ!AGN = Val(cKJAGN)
            iFeld = 39
            rsKJ!EAN = cKJEAN
            iFeld = 40
            rsKJ!MWST = cKJMwst
            iFeld = 41
            rsKJ!ekpr = dKJEkpr
            iFeld = 42
            rsKJ!vkpr = dKJVkpr
            iFeld = 43
            rsKJ!BELEGNR = gdBonNr
            
'            'ist Preiskz = 6 also Netto dann ums_ok = N
'            If Val(sPreisKz) = 6 Then
'                cUmsOK = "N"
'            End If
            
            rsKJ!UMS_OK = cUmsOK
            iFeld = 45
            sKK_art = "GZ"
            rsKJ!kk_art = sKK_art
            rsKJ!best1 = dKJBest1
            rsKJ.Update
            rsKJ.Close
            
            If gbGutscheinBeiVKversteuern = True Then
                If Val(cKJArtNr) <> 666666 Then
                    If cUmsOK <> "N" Then
                        dKdUmsatz = dKdUmsatz + dKJPreis
                    End If
                Else
                    'auch bei Gutschein
                    If cUmsOK <> "N" Then
                        dKdUmsatz = dKdUmsatz + dKJPreis
                    End If
                End If
            Else
                If Val(cKJArtNr) <> 666666 Then
                    If cUmsOK <> "N" Then
                        dKdUmsatz = dKdUmsatz + dKJPreis
                    End If
                End If
            End If
            
            
'das war mal
'            If Val(cKJArtNr) <> 666666 Then
'                If cUmsOK <> "N" Then
'                    dKdUmsatz = dKdUmsatz + dKJPreis
'                Else
'
'                End If
'            End If
            
            If Len(cLBSatz) >= 154 Then
                If Mid(cLBSatz, 154, 1) <> "N" Then
                    dKdBonus = dKdBonus + dKJPreis
                End If
            Else
                dKdBonus = dKdBonus + dKJPreis
            End If
            
        End If
        
    Next lAktSatz
    
    'storno2Bed
    If gcIdentStornoBedienerNr1 <> "" And gcIdentStornoBedienerNr2 <> "" Then
        insertStorno2 lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, CInt(gcIdentStornoBedienerNr1), CInt(gcIdentStornoBedienerNr2)
        gcIdentStornoBedienerNr1 = ""
        gcIdentStornoBedienerNr2 = ""
    End If
    
    If frmWKL68.Text1(0).Text <> "" And IsNumeric(frmWKL68.Text1(0).Text) Then 'Bar
        cGZart = "BA"
        dGeldwert = CDbl(frmWKL68.Text1(0).Text)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
    End If
    
    Dim cErzielterPreis As String
    Dim cArtNr As String
    Dim dNichtUmsatz As Double
    dNichtUmsatz = 0
    
    lAnzSatz = frmWKL20!List1.ListCount
    For lAktSatz = 0 To lAnzSatz - 1
         cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        If Len(cLBSatz) > 156 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If

        cArtNr = Mid(cLBSatz, 7, 6)
        
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        If gbGutscheinBeiVKversteuern = True Then
            If cUmsOK <> "N" Then
    
            Else
                dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
            End If
        Else
        
            If cArtNr <> "666666" Then
                If cUmsOK <> "N" Then
                Else
                    dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                End If
            End If
        End If
    Next lAktSatz
    
    
'''''
'''''    If dNichtUmsatz > 0 Then
'''''
'''''        If frmWKL68.Text1(0).Text <> "" And IsNumeric(frmWKL68.Text1(0).Text) Then 'bar
'''''
'''''            Dim dTEmpWert As Double
'''''            dTEmpWert = CDbl(frmWKL68.Text1(0).Text)
'''''
'''''            dNichtUmsatz = dNichtUmsatz - dTEmpWert
'''''        End If
'''''
'''''    End If
'''''
    
        
    
    If frmWKL68.Text1(2).Text <> "" And IsNumeric(frmWKL68.Text1(2).Text) Then  '1.Karte
        'cGZart = Right(frmWKL68.Label33(5).Caption, 4)
        cGZart = "(" & gcKreditKarte & ")"
        cGZart = Mid(cGZart, 2, 2)
        dGeldwert = CDbl(frmWKL68.Text1(2).Text)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
        insertKKZAHLTE lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
        
        If dGeldwert > 0 Then
            If dNichtUmsatz > 0 Then
                If dNichtUmsatz > dGeldwert Then
                    eintragen_AFCSTAT_NUMSKARTE dGeldwert
                Else
                    eintragen_AFCSTAT_NUMSKARTE dNichtUmsatz
                End If
                
                dNichtUmsatz = dNichtUmsatz - dGeldwert
            End If
        End If
        
    End If
    
    If frmWKL68.Text1(7).Text <> "" And IsNumeric(frmWKL68.Text1(7).Text) Then '2.Karte
    
        'cGZart = Right(frmWKL68.Label33(17).Caption, 4)
        cGZart = "(" & gcKreditKarte2 & ")"
        cGZart = Mid(cGZart, 2, 2)
        dGeldwert = CDbl(frmWKL68.Text1(7).Text)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
        insertKKZAHLTE lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
        
        If dGeldwert > 0 Then
            If dNichtUmsatz > 0 Then
                If dNichtUmsatz > dGeldwert Then
                    eintragen_AFCSTAT_NUMSKARTE dGeldwert
                Else
                    eintragen_AFCSTAT_NUMSKARTE dNichtUmsatz
                End If
            End If
        End If
        
        
        
    End If
    

    
    
    
    If gsABOPLUS_KARTE <> "" Then
        insertAboPlus_ums lKJADate, cKJAZeit, cKJBelegNr, gsABOPLUS_KARTE, gdABOPLUS_WERT
        gsABOPLUS_KARTE = ""
        gdABOPLUS_WERT = 0
    End If
    
    If frmWKL68.Text1(1).Text <> "" And IsNumeric(frmWKL68.Text1(1).Text) Then 'Gutschein
        cGZart = "GU"
        dGeldwert = CDbl(frmWKL68.Text1(1).Text)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
        
        Dim i As Integer
        If frmWKL68.List1.ListCount > 0 Then
            For i = 0 To frmWKL68.List1.ListCount - 1
                cGutsch = ""
                cGutsch = Mid(frmWKL68.List1.list(i), 1, InStr(1, frmWKL68.List1.list(i), " "))
                dGeldwert = CDbl(Mid(frmWKL68.List1.list(i), InStr(1, frmWKL68.List1.list(i), " "), Len(frmWKL68.List1.list(i)) - 1 - Len(cGutsch)))
                insertGUTZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, "EI", dGeldwert, cGutsch, CStr(lKJBediener)
            Next i
        End If
    End If
    
    If frmWKL68.Text1(3).Text <> "" And IsNumeric(frmWKL68.Text1(3).Text) Then 'Dukate
        cGZart = "DU"
        dGeldwert = CDbl(frmWKL68.Text1(3).Text)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
    End If
    
    If frmWKL68.Text1(6).Text <> "" And IsNumeric(frmWKL68.Text1(6).Text) Then 'Scheck
        cGZart = "SC"
        dGeldwert = CDbl(frmWKL68.Text1(6).Text)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
    End If
    
    If frmWKL68.Text1(5).Text <> "" And IsNumeric(frmWKL68.Text1(5).Text) Then 'EC LAST
        cGZart = "LS"
        dGeldwert = CDbl(frmWKL68.Text1(5).Text)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
        insertLASTZAHLTE lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
    End If
    
    If CLng(frmWKL68.Label333(0).Caption) <> 0 Then 'Zahlbetrag
        cGZart = "AB"
        dGeldwert = CDbl(frmWKL68.Label333(0).Caption)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
    End If
    
    If CLng(frmWKL68.Label333(3).Caption) <> 0 Then 'zurück in Bar
        cGZart = "ZB"
        dGeldwert = CDbl(frmWKL68.Label333(3).Caption)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
    End If
    
    If CLng(frmWKL68.Label1(28).Caption) <> 0 Then 'zurück als Gutschein
        cGZart = "ZG"
        dGeldwert = CDbl(frmWKL68.Label1(28).Caption)
        insertGEMZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, cGZart, dGeldwert
        
        If frmWKL68.Label1(31).Caption <> "0" Then
            cGutsch = ""
            cGutsch = frmWKL68.Label1(31).Caption
            
            insertGUTZ lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, "RG", dGeldwert, cGutsch, CStr(lKJBediener)
        End If
        
        
    End If
    
    If Val(cKJKundNr) > 0 Then
        Plus_Bonus cKJKundNr, dKdUmsatz, dKdBonus
    End If
   
    If gbbonusausjetzt = True Then
    
        Dim Bonusneu        As String
        Dim Bonusalt        As String
        Dim cDatum          As String
        Dim czeit           As String
        
        Dim cVname          As String
        Dim cNName           As String
        
        ReDim cZeilen(0 To 9) As String
    
        cSQL = "Select * from Kunden where kundnr = " & cKJKundNr
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.Edit
            rsrs!Status = "E"
            rsrs!SYNStatus = "E"
            
            If Not IsNull(rsrs!BONUS) Then
                Bonusalt = Format(rsrs!BONUS, "##000.00 EUR")
                rsrs!BONUS = rsrs!BONUS - gdbonusHerabwert
                Bonusneu = Format(rsrs!BONUS, "##000.00 EUR")
            Else
                Bonusalt = "000.00 EUR"
                rsrs!BONUS = gdbonusHerabwert * (-1)
                Bonusneu = Format(rsrs!BONUS, "##000.00 EUR")
            End If
            
            If Not IsNull(rsrs!TBONUS) Then
                rsrs!TBONUS = CDbl(rsrs!TBONUS) - gdbonusHerabwert
            Else
                rsrs!TBONUS = gdbonusHerabwert * (-1)
            End If
            
            rsrs!LASTDATE = DateValue(Now)
            rsrs!LASTTIME = TimeValue(Now)
            rsrs.Update
            
            schreibeKBProtokoll Space(10 - Len(cKJKundNr)) & cKJKundNr & " Bonus reduziert jetzt: " & Bonusneu & " vorher: " & Bonusalt
            
            cDatum = DateValue(Now)
            czeit = TimeValue(Now)
            
            cVname = lookingForKundendaten(cKJKundNr).vorname
            cNName = lookingForKundendaten(cKJKundNr).nachname
            
            cZeilen(0) = "Bonus reduziert"
            cZeilen(1) = "-----------------"
            cZeilen(2) = "KundNr:  " & cKJKundNr
            cZeilen(3) = "Vorname: " & cVname
            cZeilen(4) = "Name:    " & cNName
            cZeilen(5) = "vorher:  " & Bonusalt
            cZeilen(6) = "jetzt:   " & Bonusneu
            cZeilen(7) = "Datum:   " & cDatum
            cZeilen(8) = "Zeit:    " & czeit
            
            DruckeArbeitszeitBelegWK20d cZeilen(), 8
            
            If gb2BONUSMESS Then
                DruckeArbeitszeitBelegWK20d cZeilen(), 8
            End If
        End If
        
        rsrs.Close: Set rsrs = Nothing

    End If
    
    gdbonusHerabwert = 0
    gbbonusausjetzt = False
    gbbonusHerab = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "InsertAFCBuchModul20For68"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN/M20 ist ein Fehler aufgetreten." & Trim$(Str$(iFeld))
    
    Fehlermeldung1
    
'    Resume Next
    
End Sub
Public Sub eintragen_AFCSTAT_NUMSKARTE(dWert As Double)
On Error GoTo LOKAL_ERROR

    If dWert = 0 Then
        Exit Sub
    End If
    
    updateafcstat "NUMSKARTE", dWert, gcKasNum


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "eintragen_AFCSTAT_NUMSKARTE"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ABINFEEDB(cART As String, cMenge As String, dTranspack As Double, Optional sVonFil As String = "")
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from Feedb where artnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!edate = DateValue(Now)
    rsGZ!ezeit = TimeValue(Now)
    rsGZ!artnr = cART
    rsGZ!Menge = cMenge
    rsGZ!FILIALE = gcFilNr
    rsGZ!TRANSPACK = dTranspack
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing
    
    
    
    cSQL = "Select * from Feedb_TRANS where artnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!edate = DateValue(Now)
    rsGZ!ezeit = TimeValue(Now)
    rsGZ!artnr = cART
    rsGZ!Menge = cMenge
    rsGZ!AN_FILIALE = gcFilNr
    rsGZ!VON_FILIALE = sVonFil
    rsGZ!TRANSPACK = dTranspack
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ABINFEEDB"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ABINFEEDBF(cART As String, cMenge As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset
    
    If Val(cMenge) > 99 Then
        Exit Sub
    End If
    

    cSQL = "Select * from FeedbF where artnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!edate = DateValue(Now)
    rsGZ!ezeit = TimeValue(Now)
    rsGZ!artnr = cART
    rsGZ!Menge = cMenge
    rsGZ!ZIELFILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing
    
    'aus der eigenen Zunter löschen
    
    cSQL = "UPDATE ZUNTER set Menge = Menge - " & CInt(cMenge) & " "
    cSQL = cSQL & " where ZUNTER.filiale = " & gcFilNr & " "
    cSQL = cSQL & " and ZUNTER.artnr = " & cART
    cSQL = cSQL & " and ZUNTER.Menge - " & CInt(cMenge) & " >= 0 "
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ABINFEEDBF"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertGEMZ(lDat As Long, czeit As String, cBon As String, cKass As String, cKKart As String, dMoney As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from GEMZ where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!kasnum = cKass
    rsGZ!GELDWERT = dMoney
    rsGZ!kk_art = cKKart
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertGEMZ"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN GZ ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub insert_Gemischte_Zahlung(lDat As Long, czeit As String, dBon As Double, cKass As String, cThema As String, dMoney As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    
    cSQL = "Insert into Gemischte_Z "
    cSQL = cSQL & " (Wert "
    cSQL = cSQL & ", Thema "
    cSQL = cSQL & ", ADATE "
    cSQL = cSQL & ", Azeit "
    cSQL = cSQL & ", Filiale "
    cSQL = cSQL & ", Kasnum "
    cSQL = cSQL & ", ZBONNR "
    cSQL = cSQL & ", BELEGNR "
    cSQL = cSQL & " ) values ("
    cSQL = cSQL & "  '" & dMoney & "'  "
    cSQL = cSQL & ", '" & cThema & "'  "
    cSQL = cSQL & ", " & Trim$(Str$(lDat)) & " "
    cSQL = cSQL & ", '" & czeit & "' "
    cSQL = cSQL & ", " & gcFilNr & " "
    cSQL = cSQL & ", " & cKass & " "
    cSQL = cSQL & ", 0 "
    cSQL = cSQL & ", " & dBon & "  "
    cSQL = cSQL & " )"
    gdBase.Execute cSQL, dbFailOnError
    
    
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insert_Gemischte_Zahlung"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN GZ ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub insertStorno2(lDat As Long, czeit As String, cBon As String, cKass As String, ibednu1 As Integer, ibednu2 As Integer)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from STORNO2 where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!kasnum = cKass
    rsGZ!bednu1 = ibednu1
    rsGZ!bednu2 = ibednu2
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertStorno2"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertArtikelDetail(lDat As Long, czeit As String, cKass As String, ibednu As Integer, lartnr, sSpalte, sWert)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from ArtDet "
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!artnr = lartnr
    rsGZ!spalte = sSpalte
    rsGZ!Wert = sWert
    rsGZ!kasnum = cKass
    rsGZ!BEDNU = ibednu
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertArtikelDetail"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertArtikelMDH(lDat As Long, czeit As String, ibednu As Integer, lartnr, lMDHDAT As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset
    
    
    
    cSQL = "Delete from ArtMDH where artnr = " & lartnr & " and MDHDAT = " & Trim$(Str$(lMDHDAT)) & ""
    gdBase.Execute cSQL, dbFailOnError
    
    
    

    cSQL = "Select * from ArtMDH "
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!artnr = lartnr
    rsGZ!MDHDAT = lMDHDAT
    rsGZ!BEDNU = ibednu
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertArtikelMDH"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertKASSBEDP(lDat As Long, czeit As String, cKass As String, ibednu As Integer)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from KASSBEDP where KASNUM = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!kasnum = cKass
    rsGZ!BEDNU = ibednu
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertKASSBEDP"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertABSCHOPF(lDat As Long, czeit As String, cKass As String, ibednu1 As Integer, dMoney As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from ABSCHOPF"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!kasnum = cKass
    rsGZ!BEDNU = ibednu1
    rsGZ!GELDWERT = dMoney
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertABSCHOPF"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertDukatenbestand(lDat As Long, czeit As String, cKass As String, ibednu1 As Integer, dStueck As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from DUKATENB"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!kasnum = cKass
    rsGZ!BEDNU = ibednu1
    rsGZ!DUBESTAND = dStueck
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertDukatenbestand"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertKreditZA(lDat As Long, czeit As String, lKreditdat As Long, ibednu1 As Integer, cKKart As String, lKUNDNR As Long, lartnr As Long, iMenge As Integer, cBon As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from KreditZA"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!KREAdate = lKreditdat
    rsGZ!BEDNU = ibednu1
    rsGZ!Kundnr = lKUNDNR
    rsGZ!artnr = lartnr
    rsGZ!Menge = iMenge
    rsGZ!kk_art = cKKart
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertKreditZA"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Function insert_BONUS_SYS_Back_Max(dMoney As Double) As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim cDatum As String
    cDatum = Fix(Now)
    
    Dim czeit As String
    czeit = Format$(Now, "HH:MM:SS")
    
    insert_BONUS_SYS_Back_Max = 0
    
    
    cSQL = "Select max(BONUS_NR) as Maxi from BONUS_SYS"
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            insert_BONUS_SYS_Back_Max = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    insert_BONUS_SYS_Back_Max = insert_BONUS_SYS_Back_Max + 1
    
    
    cSQL = "Insert into BONUS_SYS "
    cSQL = cSQL & " (BONUS_NR "
    cSQL = cSQL & ", BONUS_BETRAG "
    cSQL = cSQL & ", BONUS_AUSGABEDAT "
    cSQL = cSQL & ", BONUS_AUSGABEZEIT "
    cSQL = cSQL & ", BONUS_EINLDAT "
    cSQL = cSQL & ", BONUS_EINLZEIT "
    cSQL = cSQL & ", BONUS_AUSGABEFIL "
    cSQL = cSQL & ", SENDOK "
    cSQL = cSQL & " ) values ("
    cSQL = cSQL & " " & insert_BONUS_SYS_Back_Max & " "
    cSQL = cSQL & ", '" & dMoney & "'  "
    cSQL = cSQL & ", '" & cDatum & "' "
    cSQL = cSQL & ", '" & czeit & "' "
    cSQL = cSQL & ", null "
    cSQL = cSQL & ", null "
    cSQL = cSQL & ", " & gcFilNr & " "
    cSQL = cSQL & ", False "
    cSQL = cSQL & " )"
    gdBase.Execute cSQL, dbFailOnError

    

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insert_BONUS_SYS_Back_Max"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Einloesen_BONUS_SYS(lBO_NR As Long, lFil As Long) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim cDatum As String
    cDatum = Fix(Now)
    
    Dim czeit As String
    czeit = Format$(Now, "HH:MM:SS")
    
    Einloesen_BONUS_SYS = ""

    cSQL = "Select BONUS_BETRAG from BONUS_SYS where BONUS_NR = " & lBO_NR & ""
    cSQL = cSQL & " and BONUS_EINLZEIT is null and  BONUS_EINLDAT is null "
    If gsWWBonusGDAUER > 0 Then
        'auch die Gültigkeit überprüfen
        cSQL = cSQL & " and BONUS_AUSGABEDAT >= " & CLng(DateValue(Now) - gsWWBonusGDAUER)
    End If
    
    cSQL = cSQL & " and BONUS_AUSGABEFIL = " & lFil & ""
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!BONUS_BETRAG) Then
            Einloesen_BONUS_SYS = rsrs!BONUS_BETRAG
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Einloesen_BONUS_SYS <> "" Then
    
        cSQL = "Update BONUS_SYS set BONUS_EINLDAT = '" & cDatum & "' "
        cSQL = cSQL & ", BONUS_EINLZEIT = '" & czeit & "' "
        cSQL = cSQL & ", sendok = False "
        cSQL = cSQL & " where BONUS_NR = " & lBO_NR & ""
        gdBase.Execute cSQL
    Else
    
        Einloesen_BONUS_SYS = "Dieser Bonus kann nicht ausgezahlt werden." & vbCrLf & vbCrLf
        
        
        'es gibt hier 2 Gründe
        
        Dim bErsterGrund As Boolean
        bErsterGrund = False
        
        '1. Grund Gültigkeit überschritten
        
        cSQL = "Select BONUS_EINLDAT, BONUS_AUSGABEDAT from BONUS_SYS where BONUS_NR = " & lBO_NR & ""
        FnOpenrecordset rsrs, cSQL, 1, gdBase
        
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!BONUS_EINLDAT) Then
            
            Else
                bErsterGrund = True
                
                If Not IsNull(rsrs!BONUS_AUSGABEDAT) Then
                    Einloesen_BONUS_SYS = Einloesen_BONUS_SYS & "Die Gültigkeit von " & gsWWBonusGDAUER & " Tagen ist am " & Format(DateValue(rsrs!BONUS_AUSGABEDAT) + gsWWBonusGDAUER, "DD.MM.YYYY") & " abgelaufen."
                End If
            End If
            
            
            
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
        
        
        
        
        If bErsterGrund = False Then
            '2. Grund schon eingelöst
        
            cSQL = "Select BONUS_BETRAG, BONUS_EINLDAT, BONUS_EINLZEIT from BONUS_SYS where BONUS_NR = " & lBO_NR & ""
            FnOpenrecordset rsrs, cSQL, 1, gdBase
            
            If Not rsrs.EOF Then
            
                If Not IsNull(rsrs!BONUS_BETRAG) Then
                    Einloesen_BONUS_SYS = Einloesen_BONUS_SYS & "Dieser Betrag von: " & Format(rsrs!BONUS_BETRAG, "####0.00") & " wurde"
                End If
                
                If Not IsNull(rsrs!BONUS_EINLDAT) Then
                    Einloesen_BONUS_SYS = Einloesen_BONUS_SYS & " am: " & rsrs!BONUS_EINLDAT & " "
                End If
                
                If Not IsNull(rsrs!BONUS_EINLZEIT) Then
                    Einloesen_BONUS_SYS = Einloesen_BONUS_SYS & " um: " & rsrs!BONUS_EINLZEIT & " "
                End If
                
                Einloesen_BONUS_SYS = Einloesen_BONUS_SYS & " eingelöst."
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
    End If


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Einloesen_BONUS_SYS"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub insertKKZAHLTE(lDat As Long, czeit As String, cBon As String, cKass As String, cKKart As String, dMoney As Double)
On Error GoTo LOKAL_ERROR
 
    If lDat = 0 Or Trim(czeit) = "" Or Trim(cBon) = "" Then
     
      Exit Sub
      
    End If
    
 
    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from KKZAHLTE where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!kasnum = cKass
    rsGZ!GELDWERT = dMoney
    rsGZ!kk_art = cKKart
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "insertKKZAHLTE"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN KK ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertAboPlus_ums(lDat As Long, czeit As String, cBon As String, cKKart As String, dMoney As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsGZ    As Recordset
    Dim lPFNR   As Long
    
    cSQL = "Select distinct(PFNR) from ABOPLUS"
    Set rsGZ = gdBase.OpenRecordset(cSQL)
    If Not rsGZ.EOF Then
        If Not IsNull(rsGZ!PFNR) Then
            lPFNR = rsGZ!PFNR
        End If
    End If
    rsGZ.Close: Set rsGZ = Nothing
    
    cSQL = "Select * from ABOPLUS_UMS where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!GELDWERT = dMoney
    rsGZ!ABOPLUSKARTE = cKKart
    rsGZ!FILIALE = gcFilNr
    rsGZ!PFNR = lPFNR
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertAboPlus_ums"
    Fehler.gsFehlertext = "Im Programteil AboPlusKarte ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub insertNichtUmsBar(lDat As Long, czeit As String, cBon As String, cKass As String, dMoney As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset
    
    If lDat = 0 Then
        Exit Sub
    End If

    cSQL = "Select * from NichtUmsBar where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew

    rsGZ!Datum = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!kasnum = cKass
    rsGZ!Betrag = dMoney
    rsGZ!art = ""
    rsGZ!FILIALE = gcFilNr
    rsGZ!BEDNU = 0
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertNichtUmsBar"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN BAR ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub




Public Sub insertLASTZAHLTE(lDat As Long, czeit As String, cBon As String, cKass As String, cKKart As String, dMoney As Double)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset
    
    If lDat = 0 Then
        Exit Sub
    End If

    cSQL = "Select * from LASTZAHLTE where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!kasnum = cKass
    rsGZ!GELDWERT = dMoney
    rsGZ!kk_art = cKKart
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertLASTZAHLTE"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN EC LAST ist ein Fehler aufgetreten." & lDat & " " & czeit & " " & cBon & " " & cKass & " " & dMoney & " " & cKKart & " " & gcFilNr & " "
    
    Fehlermeldung1
End Sub
Public Sub insertGUTZ(lDat As Long, czeit As String, cBon As String, cKass As String, cKKart As String, dMoney As Double, cGutsch As String, cBed As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from GUTZ where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!kasnum = cKass
    rsGZ!GELDWERT = dMoney
    rsGZ!art = cKKart
    rsGZ!gutschnr = cGutsch
    rsGZ!BEDNU = cBed
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertGUTZ"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN GZ ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertGUTschHIS(lDat As Long, czeit As String, cBon As String, cKass As String, cKKart As String, dMoney As Double, cGutsch As String, cBed As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from GUHIS where belegnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!BELEGNR = cBon
    rsGZ!kasnum = cKass
'    rsGZ!GeldWERT = dMoney
    rsGZ!art = cKKart
    rsGZ!gutschnro = cGutsch
    rsGZ!BEDNU = cBed
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertGUTschHIS"
    Fehler.gsFehlertext = "Im Programteil BEZAHLEN GZ ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermWertforTSS(sMWST As String, sUms As String, sTab As String) As Double
On Error GoTo LOKAL_ERROR

    ermWertforTSS = 0
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    

    sSQL = "Select sum(APREIS) as Wert from " & sTab & " where "
    sSQL = sSQL & "  AMWSK = '" & sMWST & "'"
    sSQL = sSQL & " and UMS_OK = '" & sUms & "'"
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Wert) Then
            ermWertforTSS = rsrs!Wert
        End If
    End If
    rsrs.Close


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermWertforTSS"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub UpdateAFCStatGutscheinModul20(dZuZahlen As Double, dEinrGutsch As Double, dRestZhlg As Double, dRestGutschein As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim lStornoAnz      As Long
    Dim lDatum          As Long
    Dim lAktSatz        As Long
    Dim lAnzSatz        As Long
    
    Dim cArtNr          As String
    Dim cUmsOK          As String
    Dim cSQL            As String
    Dim cErzielterPreis As String
    Dim ctmp            As String
    Dim cNormal         As String
    Dim cPosSumme       As String
    Dim cArtRabatt      As String
    Dim cLBSatz         As String
    
    Dim dUmsatz         As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dEchterUmsatz   As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dFalscherUmsatz As Double     'Summe des Verkaufs ohne Gutscheine
    Dim dBarAuszahlung  As Double
    Dim dUmsatz2        As Double     'Summe des Verkaufs inkl. Gutscheine
    Dim dSPreisAnz      As Double     'Anzahl Positionen mit Sonderpreis
    Dim dSPreisGes      As Double     'Summe aller Positionen mit Sonderpreis
    Dim dKundenZahl     As Double     'Konstante 1
    Dim dArtRabAnz      As Double     'Anzahl Positionen mit Artikelrabatt
    Dim dArtRabGes      As Double     'Summe des gegebenen Artikelrabatts
    Dim dGesRabAnz      As Double     'Anzahl Positionen mit Gesamtrabatt
    Dim dGesRabGes      As Double     'Summe des gegebenen Gesamtrabatts
    Dim dWertGutschein  As Double
    Dim dRestgutsch     As Double
    Dim dUmsatzGutsch   As Double
    Dim dStornoWert     As Double
    Dim dZhlgGutsch     As Double
    Dim rsrs            As Recordset
   
    
    
    
    
    lDatum = Fix(Now)
    
    'feste Werte setzen
    dKundenZahl = 0
    
    '*******************************************
    '* Was hat der Kunde insgesamt zu zahlen?
    '*******************************************
    dUmsatz = dZuZahlen
    
    '*******************************************
    '* Was zahlt der Kunde mittels Gutscheinen?
    '*******************************************
    dZhlgGutsch = dEinrGutsch
    
    dEchterUmsatz = 0
    dWertGutschein = 0
    
    '*******************************************
    '* Untersuche jeden einzelnen Artikel
    '*******************************************
    
    lAnzSatz = frmWKL20!List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        
        cArtNr = Mid(cLBSatz, 7, 6)
        
'        ctmp = Mid(cLBSatz, 148, 3)
'        ctmp = Trim$(ctmp)
'        lKJBediener = Val(ctmp)
        
        '*******************************************
        '* Lies Kennzeichen Umsatzrelevanz
        '*******************************************
        If Len(cLBSatz) > 155 Then
            cUmsOK = Mid(cLBSatz, 156, 1)
        Else
            cUmsOK = "J"
        End If
        If cUmsOK <> "J" And cUmsOK <> "N" Then
            cUmsOK = "J"
        End If
                
        '*******************************************
        '* Lies regulären Stückpreis
        '*******************************************
        cNormal = Mid(cLBSatz, 128, 9)
        cNormal = Trim$(cNormal)
        cNormal = fnMoveComma2Point$(cNormal)
        
        '*******************************************
        '* Lies Stückpreis, zu dem verkauft wurde
        '*******************************************
        ctmp = Mid(cLBSatz, 74, 9)
        ctmp = Trim$(ctmp)
        ctmp = fnMoveComma2Point$(ctmp)
        
        '*******************************************
        '* Lies Gesamtpreis der Position
        '*******************************************
        cPosSumme = Mid(cLBSatz, 94, 9)
        cPosSumme = Trim$(cPosSumme)
        cPosSumme = fnMoveComma2Point$(cPosSumme)
        
        '*******************************************
        '* Lies Artikelrabatt der Position
        '*******************************************
        cArtRabatt = Mid(cLBSatz, 124, 3)
        cArtRabatt = Trim$(cArtRabatt)
        cArtRabatt = fnMoveComma2Point$(cArtRabatt)
        
        '**********************************************
        '* Lies den echten Verkaufspreis der Position
        '**********************************************
        cErzielterPreis = Mid(cLBSatz, 60, 9)
        cErzielterPreis = Trim$(cErzielterPreis)
        cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
        
        '**********************************************
        '* Ermittle die echte Umsatzsumme
        '* - keine Gutscheine
        '* - kein nicht umsatzrelevanten Artikel
        '**********************************************
        If cArtNr <> "666666" Then
            If cUmsOK <> "N" Then
                dEchterUmsatz = dEchterUmsatz + Val(cErzielterPreis)
                dKundenZahl = 1
            End If
        Else
            '******************************************************
            '* Wenn Gutschein, dann summiere verkaufte Gutscheine
            '******************************************************
            dWertGutschein = dWertGutschein + Val(cErzielterPreis)
        End If
        
        '**********************************************
        '* Wenn regulärer Stückpreis und Stückpreis
        '* des Verkaufes abweichen, Zähler für
        '* Sonderpreis um 1 heraufsetzen und die
        '* Sonderpreissumme erhöhen
        '**********************************************
        '//2002
        If Val(ctmp) <> Val(cNormal) Then
            If cNormal = 0 Then
                dSPreisAnz = 0
                dSPreisGes = 0
            Else
                dSPreisAnz = dSPreisAnz + 1
                dSPreisGes = dSPreisGes + Val(cPosSumme)
            End If
        End If
        
        '*******************************************
        '* Wenn Artikelrabatt gewährt wurde,
        '* Zähler für Artikelrabatt um 1 heraufsetzen
        '* und ArtRabattsumme erhöhen
        '*******************************************
        If Val(cArtRabatt) <> 0 Then
            ctmp = Mid(cLBSatz, 84, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
            dArtRabAnz = dArtRabAnz + 1
            dArtRabGes = dArtRabGes + Val(ctmp)
        ElseIf frmWKL20!Label2(3).Visible Then
            dGesRabAnz = dGesRabAnz + 1
        End If
    
        '*******************************************
        '* Wenn erzielter Preis < 0, dann
        '* Zähler für Storno um 1 heraufsetzen
        '* und Stornosumme erhöhen
        '*******************************************
        If Val(cErzielterPreis) < 0 Then
            If IstArtikelnichtStornierfähig(cArtNr) = False Then
                dStornoWert = dStornoWert + Val(cErzielterPreis)
                lStornoAnz = lStornoAnz + 1
            End If
        End If
    
    Next lAktSatz
    
    

    
    '**************************************************************
    '* Die Differenz zwischen der zu zahlenden Summe und dem
    '* echten Umsatz ist der falsche Umsatz (Gutschein-Verkäufe,
    '* Verkäufe von nicht umsatzrelevanten Artikeln)
    '**************************************************************
    
    dFalscherUmsatz = dZuZahlen - dEchterUmsatz
    
    '**************************************************************
    '* Wieviel hat der Kunde sich in Bar auszahlen lassen?
    '**************************************************************
    
    dBarAuszahlung = dZhlgGutsch - dEchterUmsatz - dFalscherUmsatz - dRestGutschein
    If dBarAuszahlung < 0 Then
        dBarAuszahlung = 0
    End If
    
    '**************************************************************
    '* Wurden mit einem Gutschein neue Gutscheine gekauft?
    '* Wenn ja, dann reduziert sich der Umsatz aus Gutscheinen
    '* um den Wert der verkauften Gutscheine (bis max. auf 0)
    '**************************************************************
    
    
    If dZhlgGutsch > dFalscherUmsatz Then
        dUmsatzGutsch = dZhlgGutsch - dFalscherUmsatz - dRestGutschein - dBarAuszahlung
    Else
        dUmsatzGutsch = dZhlgGutsch - dRestGutschein - dBarAuszahlung
    End If

    
    
    If dEchterUmsatz > dUmsatzGutsch Then
        dUmsatz = dEchterUmsatz - dUmsatzGutsch
    Else
        dUmsatz = 0                                 'Umsatz aus sonstigen Zahlungen (außer Gutschein)
    End If
    
    If frmWKL20!Label2(3).Visible Then
        dGesRabGes = fnHoleGesamtRabattModul20#()
    End If
    
    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    Select Case giZahlArt
        
        Case Is = 1   '//zahlung mit Gutschein + EC-Lastschrift
            If Not IsNull(rsrs!UMS_LAST) Then
                rsrs!UMS_LAST = rsrs!UMS_LAST + gdGutLastRest
            Else
                rsrs!UMS_LAST = gdGutLastRest
            End If
            
        Case Is = 8         'BAR
            '**************************************************************
            '* Wieviel Umsatz hat der Kunde in bar gemacht?
            '**************************************************************
            'erzielter Umsatz über Bargeld
            If Not IsNull(rsrs!UMS_BAR) Then
                rsrs!UMS_BAR = rsrs!UMS_BAR + dUmsatz
            Else
                rsrs!UMS_BAR = dUmsatz
            End If
            


            'Zugang Bargeld in Kasse
            
            If Not IsNull(rsrs!BARVERKAUF) Then
            
'                If gbGutschOverBar Then
'                    rsrs!BARVERKAUF = rsrs!BARVERKAUF + 0
'                Else
                    rsrs!BARVERKAUF = rsrs!BARVERKAUF + dRestZhlg
'                End If
                
            Else
'                If gbGutschOverBar Then
'                    rsrs!BARVERKAUF = 0
'                Else
                    rsrs!BARVERKAUF = dRestZhlg
'                End If
            End If

            gbGutschOverBar = False
            
            'Teile der Restzahlung fließen in neuen Gutschein

            If dRestZhlg > dUmsatz Then
                If Not IsNull(rsrs!GUTSCHBAR) Then
                    rsrs!GUTSCHBAR = rsrs!GUTSCHBAR + Format(dRestZhlg - dUmsatz, "######0.00") 'Hier ist Format() neu 21.10.13
                Else
                    rsrs!GUTSCHBAR = Format(dRestZhlg - dUmsatz, "######0.00") 'Hier ist Format() neu 21.10.13
                End If
            
            End If
        
        Case Is = 5         'KREDIT
            '**************************************************************
            '* Wieviel Umsatz hat der Kunde über Kredit gemacht?
            '**************************************************************
            If Not IsNull(rsrs!UMS_Kred) Then
                rsrs!UMS_Kred = rsrs!UMS_Kred + dUmsatz
            Else
                rsrs!UMS_Kred = dUmsatz
            End If
            'Teile der Restzahlung fließen in neuen Gutschein
            If dRestZhlg > dUmsatz Then
                If Not IsNull(rsrs!GUTSCHKRE) Then
                    rsrs!GUTSCHKRE = rsrs!GUTSCHKRE + (dRestZhlg - dUmsatz)
                Else
                    rsrs!GUTSCHKRE = (dRestZhlg - dUmsatz)
                End If
            End If
            
        Case Is = 6         'SCHECK
            '**************************************************************
            '* Wieviel Umsatz hat der Kunde mit Schecks gemacht?
            '**************************************************************
            If Not IsNull(rsrs!UMS_SCHECK) Then
                rsrs!UMS_SCHECK = rsrs!UMS_SCHECK + dUmsatz
            Else
                rsrs!UMS_SCHECK = dUmsatz
            End If
            'Zugang Schecks in Kasse
            If Not IsNull(rsrs!SCHVERKAUF) Then
                rsrs!SCHVERKAUF = rsrs!SCHVERKAUF + dRestZhlg
            Else
                rsrs!SCHVERKAUF = dRestZhlg
            End If
            If Not IsNull(rsrs!ANZSCHECK) Then
                rsrs!ANZSCHECK = rsrs!ANZSCHECK + dKundenZahl
            Else
                rsrs!ANZSCHECK = dKundenZahl
            End If
            'Teile der Restzahlung fließen in neuen Gutschein
            If dRestZhlg > dUmsatz Then
                If Not IsNull(rsrs!GUTSCHSCH) Then
                    rsrs!GUTSCHSCH = rsrs!GUTSCHSCH + (dRestZhlg - dUmsatz)
                Else
                    rsrs!GUTSCHSCH = (dRestZhlg - dUmsatz)
                End If
            End If
        Case Is = 17            'KREDITKARTE
            '**************************************************************
            '* Wieviel Umsatz hat der Kunde mit Kreditkarten gemacht?
            '**************************************************************
            
            '08.06.04
            
            If Not IsNull(rsrs!UMS_KARTE) Then
                rsrs!UMS_KARTE = rsrs!UMS_KARTE + dUmsatz
                schreibeProtokollUNITXT CStr(dUmsatz), "Kartenzahlung"
            Else
                rsrs!UMS_KARTE = dUmsatz
                schreibeProtokollUNITXT CStr(dUmsatz), "Kartenzahlung"
            End If
            




        Case Is = 46                'EINZAHLUNG
            '**************************************************************
            '* Wieviel Geld ist eingezahlt worden (Bargeld-Kasseneinlage)
            '**************************************************************
            If Not IsNull(rsrs!EINZAHLUNG) Then
                rsrs!EINZAHLUNG = rsrs!EINZAHLUNG + dUmsatz2
            Else
                rsrs!EINZAHLUNG = dUmsatz2
            End If
            
        Case Is = 45                'AUSZAHLUNG
            '**************************************************************
            '* Wieviel Geld ist ausgezahlt worden (Bargeld-Kassenentnahme)
            '**************************************************************
            If Not IsNull(rsrs!AUSZAHLUNG) Then
                rsrs!AUSZAHLUNG = rsrs!AUSZAHLUNG + dUmsatz2
            Else
                rsrs!AUSZAHLUNG = dUmsatz2
            End If
        
    End Select
    
    '**************************************************************
    '* Datum und Kassennummer des Verbuchens schreiben
    '**************************************************************
    rsrs!ADATE = lDatum
    rsrs!kasnum = Val(gcKasNum)

    '**************************************************************
    '* Betrag der eingereichten Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!EINRGUTSCH) Then
        rsrs!EINRGUTSCH = rsrs!EINRGUTSCH + dZhlgGutsch
    Else
        rsrs!EINRGUTSCH = dZhlgGutsch
    End If


    '**************************************************************
    '* Betrag der generierten Rest-Gutscheine verbuchen
    '**************************************************************
    If Not IsNull(rsrs!RESTGUTSCH) Then
        rsrs!RESTGUTSCH = rsrs!RESTGUTSCH + dRestGutschein
    Else
        rsrs!RESTGUTSCH = dRestGutschein
    End If

    '**************************************************************
    '* Betrag der Gutschein-AUszahlung verbuchen
    '**************************************************************
    If Not IsNull(rsrs!AUSZGUTSCH) Then
        rsrs!AUSZGUTSCH = Format(rsrs!AUSZGUTSCH + dBarAuszahlung, "######0.00") 'Hier ist Format() neu 21.10.13
    Else
        rsrs!AUSZGUTSCH = Format(dBarAuszahlung, "######0.00") 'Hier ist Format() neu 21.10.13
    End If

    '**************************************************************
    '* Betrag des Umsatzes durch Gutschein-Einreichungen verbuchen
    '**************************************************************
    
    If Not IsNull(rsrs!ZHLGGUTSCH) Then
        rsrs!ZHLGGUTSCH = rsrs!ZHLGGUTSCH + dUmsatzGutsch
    Else
        rsrs!ZHLGGUTSCH = dUmsatzGutsch
    End If
    
    '**************************************************************
    '* Betrag des Gutschein-Verkäufe verbuchen
    '**************************************************************
    If Not IsNull(rsrs!GUTSCHEIN) Then
        rsrs!GUTSCHEIN = rsrs!GUTSCHEIN + dWertGutschein
    Else
        rsrs!GUTSCHEIN = dWertGutschein
    End If
    
    '**************************************************************
    '* Sonderpreise verbuchen
    '**************************************************************
    
    If Not IsNull(rsrs!SPREIS_ANZ) Then
        rsrs!SPREIS_ANZ = rsrs!SPREIS_ANZ + dSPreisAnz
    Else
        rsrs!SPREIS_ANZ = dSPreisAnz
    End If
    If Not IsNull(rsrs!SPREIS_GES) Then
        rsrs!SPREIS_GES = rsrs!SPREIS_GES + dSPreisGes
    Else
        rsrs!SPREIS_GES = dSPreisGes
    End If
    
    '**************************************************************
    '* Kundenzahl schreiben
    '**************************************************************
    
    If Not IsNull(rsrs!Kundenzahl) Then
        rsrs!Kundenzahl = rsrs!Kundenzahl + dKundenZahl
    Else
        rsrs!Kundenzahl = dKundenZahl
    End If
    
    '**************************************************************
    '* Artikelrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!ARTRAB_ANZ) Then
        rsrs!ARTRAB_ANZ = rsrs!ARTRAB_ANZ + dArtRabAnz
    Else
        rsrs!ARTRAB_ANZ = dArtRabAnz
    End If
    If Not IsNull(rsrs!ARTRAB_GES) Then
        rsrs!ARTRAB_GES = rsrs!ARTRAB_GES + dArtRabGes
    Else
        rsrs!ARTRAB_GES = dArtRabGes
    End If
    
    '**************************************************************
    '* Gesamtrabatte schreiben
    '**************************************************************
    If Not IsNull(rsrs!GESRAB_ANZ) Then
        rsrs!GESRAB_ANZ = rsrs!GESRAB_ANZ + dGesRabAnz
    Else
        rsrs!GESRAB_ANZ = dGesRabAnz
    End If
    If Not IsNull(rsrs!GESRAB_GES) Then
        rsrs!GESRAB_GES = rsrs!GESRAB_GES + dGesRabGes
    Else
        rsrs!GESRAB_GES = dGesRabGes
    End If
    
    
    '**************************************************************
    '* Bonnummer schreiben
    '**************************************************************

    If gdBonNr = 0 Then
        HoleNeueBonNrWKL20_NEU 'bonnr wird gleich eingetragen
    End If

    If Not IsNull(rsrs!BELEGNR) Then
        If gdBonNr < CLng(rsrs!BELEGNR) Then
        
        
            
        Else
            rsrs!BELEGNR = gdBonNr
        End If
    Else
        rsrs!BELEGNR = gdBonNr
    End If
    
    '**************************************************************
    '* Stornos schreiben
    '**************************************************************
    If Not IsNull(rsrs!STORNO_GES) Then
        rsrs!STORNO_GES = rsrs!STORNO_GES + dStornoWert
    Else
        rsrs!STORNO_GES = dStornoWert
    End If
    
    If Not IsNull(rsrs!STORNO_ANZ) Then
        rsrs!STORNO_ANZ = rsrs!STORNO_ANZ + lStornoAnz
    Else
        rsrs!STORNO_ANZ = lStornoAnz
    End If
    
    
    '**************************************************************
    '* Schreibvorgang durchführen
    '**************************************************************
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "UpdateAFCStatGutscheinModul20"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub

Public Function fnHoleGesamtRabattModul20#()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim cLBSatz As String
    Dim ctmp As String
    Dim dWert As Double
    Dim dAnz As Double
    Dim dLPreis As Double
    Dim dTPreis As Double
    Dim dVPreis As Double
    Dim dErmBetrag As Double
    
    lAnzSatz = frmWKL20!List1.ListCount
    
    If lAnzSatz = 0 Then
        fnHoleGesamtRabattModul20# = 0
        Exit Function
    End If
    
    dErmBetrag = 0
    
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        ctmp = Mid(cLBSatz, 124, 3)
        dWert = Val(ctmp)
        If dWert = 0 Then
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If
            ctmp = Trim$(ctmp)
            dAnz = Val(ctmp)
            
            ctmp = Mid(cLBSatz, 74, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dLPreis = Val(ctmp)
            
            dTPreis = dLPreis * dAnz
            
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dVPreis = Val(ctmp)
            
            dVPreis = dTPreis - dVPreis
            
            dErmBetrag = dErmBetrag + dVPreis
            
        End If
    Next lAktSatz
    
    fnHoleGesamtRabattModul20# = dErmBetrag
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fnHoleGesamtRabattModul20"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub DruckeGutscheinBonModul20(cLBSatz As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cDaten As String
    Dim cEscapeSequenz As String
    Dim aDeviceName As String
    Dim lAnzZeile As Long
    Dim lcount As Long
    ReDim cDruckZeile(1 To 1) As String
    Dim iLenZeile As Integer
    Dim ctmp As String
    Dim cGPreis As String
    Dim iStufe As Integer
    
    If Not gbBonDruck Then
        GoTo ENDE
    End If
    
    iStufe = 0
    
    cGPreis = Mid(cLBSatz, 60, 9)
    ctmp = Mid(cLBSatz, 24, 8)
    
    iLenZeile = 32
    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker
    
    
    '***********************************************
    'ggf. Logo auf Kassenbon bringen
    '***********************************************

    If gcBild <> "" Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If

    cEscapeSequenz = gcInit
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    

    Barcode_Gutschein Trim(ctmp)

    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    

    iStufe = 1
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(0)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        iStufe = 2
    End If
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(1)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        iStufe = 3
    End If
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(4)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        iStufe = 3
    End If
    
    '***********************************************
    'Kopfdaten 4.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(12)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    iStufe = 4
    
    '******************************************************************
    
    cDaten = "G U T S C H E I N V E R K A U F"
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    iStufe = 5
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    iStufe = 6
    '******************************************************************
        
    cDaten = "Wert Gutscheins:    " & gcWaehrung & cGPreis
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 7
    '******************************************************************
    

    cDaten = "Nummer Gutschein:" & Space(15 - Len(ctmp)) & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 8
    '******************************************************************
    
    cDaten = "Nummer der Filiale:            1"
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    ctmp = "Kasse:                         " & gcKasNum
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 9
    '******************************************************************
    ctmp = Trim$(frmWKL20!Text1(0).Text)
    ctmp = Trim$(ctmp)
    ctmp = Space$(3 - Len(ctmp)) & ctmp
    cDaten = "Bedienernummer:              " & ctmp
'    ctmp = Space$(2 - Len(ctmp)) & ctmp
'    cDaten = "Bedienernummer:               " & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 10
    '******************************************************************
    ctmp = Trim$(Str$(gdBonNr))
    ctmp = Trim$(ctmp)
    ctmp = gcKasNum & "/" & ctmp
    ctmp = Space$(10 - Len(ctmp)) & ctmp
    
    cDaten = "Belegnummer:          " & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 11
    '******************************************************************
    ctmp = Format$(Now, "DD.MM.YYYY")
    cDaten = "Datum:                " & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    iStufe = 12
    '******************************************************************
    
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    
    
    If gbBonGu2J Then
        cDaten = "Gültigkeitsdauer 4 Jahre"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        cDaten = "nach Ausstellungsdatum"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    
    
    
    
    iStufe = 13
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iStufe = 14
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iStufe = 15
    End If
    '******************************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = ""
    Else
        cDaten = gcBonText(5)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        iStufe = 15
    End If
    '******************************************************************
    
    For lcount = 1 To 9
        If lcount = 9 Then
            cEscapeSequenz = "." & vbCrLf
        Else
            cEscapeSequenz = " " & vbCrLf
        End If
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
    iStufe = 16
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
    
    gdSumme = cGPreis
    
    SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False, True
    
    Erase cDruckZeile
    iStufe = 17
    
'BON_SCHNEIDEN:

    'MsgBox "Schneide Kassenbon"

    'Kassenbon abschneiden
'    If gbAPI Then
'        aDeviceName = Printer.DeviceName
'        cEscapeSequenz = gcSchneiden
'        OpenDrawer aDeviceName, cEscapeSequenz
'    End If
    iStufe = 18
    
ENDE:

    '...und tschüß!
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "DruckeGutscheinBonModul20"
    Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten." & Trim$(Str$(iStufe))
    
    Fehlermeldung1
    
End Sub
Public Sub Barcode_Gutschein(cGutschnr As String)
On Error GoTo LOKAL_ERROR

    Dim lcount                  As Long
    Dim cZeichen                As String
    Dim cArtNr                  As String
    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    Dim lPruefZiffer            As Long

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(27) & Chr(97) & Chr(1)
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(29) & Chr(72) & Chr(2)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    'die Barcodehöhe
    
'    cEscapeSequenz = Chr(29) & Chr(104) & Chr(164)
'    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(104) & Chr(40)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(119) & Chr(3)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    
    
    If Len(cGutschnr) = 1 Then
        cGutschnr = "000000000" & cGutschnr
    ElseIf Len(cGutschnr) = 2 Then
        cGutschnr = "00000000" & cGutschnr
    ElseIf Len(cGutschnr) = 3 Then
        cGutschnr = "0000000" & cGutschnr
    ElseIf Len(cGutschnr) = 4 Then
        cGutschnr = "000000" & cGutschnr
    ElseIf Len(cGutschnr) = 5 Then
        cGutschnr = "00000" & cGutschnr
    ElseIf Len(cGutschnr) = 6 Then
        cGutschnr = "0000" & cGutschnr
    ElseIf Len(cGutschnr) = 7 Then
        cGutschnr = "000" & cGutschnr
    ElseIf Len(cGutschnr) = 8 Then
        cGutschnr = "00" & cGutschnr
    ElseIf Len(cGutschnr) = 9 Then
        cGutschnr = "0" & cGutschnr
    ElseIf Len(cGutschnr) = 10 Then
        
    End If
    
    cArtNr = "2" & "1" & cGutschnr
    
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim p5 As Integer
    Dim p6 As Integer
    Dim p7 As Integer
    Dim p8 As Integer
    Dim p9 As Integer
    Dim p10 As Integer
    Dim p11 As Integer
    Dim p12 As Integer
    Dim p13 As Integer
    
    Dim rest As Double
    Dim pz As Long
    
    
    p1 = Val(Mid(cArtNr, 1, 1)) * 1
    p2 = Val(Mid(cArtNr, 2, 1)) * 3
    p3 = Val(Mid(cArtNr, 3, 1)) * 1
    p4 = Val(Mid(cArtNr, 4, 1)) * 3
    p5 = Val(Mid(cArtNr, 5, 1)) * 1
    p6 = Val(Mid(cArtNr, 6, 1)) * 3
    p7 = Val(Mid(cArtNr, 7, 1)) * 1
    p8 = Val(Mid(cArtNr, 8, 1)) * 3
    p9 = Val(Mid(cArtNr, 9, 1)) * 1
    p10 = Val(Mid(cArtNr, 10, 1)) * 3
    p11 = Val(Mid(cArtNr, 11, 1)) * 1
    p12 = Val(Mid(cArtNr, 12, 1)) * 3
    p13 = p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 + p10 + p11 + p12
    
    rest = p13 Mod 10
    pz = 10 - rest
    If rest = 0 Then
        pz = 0
    End If
    

    
    cArtNr = cArtNr & Trim$(Str$(pz))
    
    
    
    
    cEscapeSequenz = gcBarCode & cArtNr
    OpenDrawer aDeviceName, cEscapeSequenz
    
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Barcode_Gutschein"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Barcode_BonusV3(ByRef cRabwert As String)
On Error GoTo LOKAL_ERROR

    Dim lcount                  As Long
    Dim cZeichen                As String
    Dim cArtNr                  As String
    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    Dim lPruefZiffer            As Long
    Dim cBarcode                As String

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(27) & Chr(97) & Chr(1)
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(29) & Chr(72) & Chr(2)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    'die Barcodehöhe
    
'    cEscapeSequenz = Chr(29) & Chr(104) & Chr(164)
'    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(104) & Chr(40)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(119) & Chr(3)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    
    
    If Len(cRabwert) = 1 Then
        cBarcode = "0000000" & cRabwert
    ElseIf Len(cRabwert) = 2 Then
        cBarcode = "000000" & cRabwert
    ElseIf Len(cRabwert) = 3 Then
        cBarcode = "00000" & cRabwert
    ElseIf Len(cRabwert) = 4 Then
        cBarcode = "0000" & cRabwert
    ElseIf Len(cRabwert) = 5 Then
        cBarcode = "000" & cRabwert
    ElseIf Len(cRabwert) = 6 Then
        cBarcode = "00" & cRabwert
    ElseIf Len(cRabwert) = 7 Then
        cBarcode = "0" & cRabwert
    ElseIf Len(cRabwert) = 8 Then
        
    End If
    
    cArtNr = "2771" & cBarcode
    
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim p5 As Integer
    Dim p6 As Integer
    Dim p7 As Integer
    Dim p8 As Integer
    Dim p9 As Integer
    Dim p10 As Integer
    Dim p11 As Integer
    Dim p12 As Integer
    Dim p13 As Integer
    
    Dim rest As Double
    Dim pz As Long
    
    
    p1 = Val(Mid(cArtNr, 1, 1)) * 1
    p2 = Val(Mid(cArtNr, 2, 1)) * 3
    p3 = Val(Mid(cArtNr, 3, 1)) * 1
    p4 = Val(Mid(cArtNr, 4, 1)) * 3
    p5 = Val(Mid(cArtNr, 5, 1)) * 1
    p6 = Val(Mid(cArtNr, 6, 1)) * 3
    p7 = Val(Mid(cArtNr, 7, 1)) * 1
    p8 = Val(Mid(cArtNr, 8, 1)) * 3
    p9 = Val(Mid(cArtNr, 9, 1)) * 1
    p10 = Val(Mid(cArtNr, 10, 1)) * 3
    p11 = Val(Mid(cArtNr, 11, 1)) * 1
    p12 = Val(Mid(cArtNr, 12, 1)) * 3
    p13 = p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 + p10 + p11 + p12
    
    rest = p13 Mod 10
    pz = 10 - rest
    If rest = 0 Then
        pz = 0
    End If
    
    cArtNr = cArtNr & Trim$(Str$(pz))
    
    cEscapeSequenz = gcBarCode & cArtNr
    OpenDrawer aDeviceName, cEscapeSequenz
    
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
'
'    cEscapeSequenz = vbCrLf
'    OpenDrawer aDeviceName, cEscapeSequenz
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Barcode_BonusV3"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub OpenDrawerViaComPortModul20()
    On Error GoTo LOKAL_ERROR
    
    frmWKL20!MSComm2.CommPort = Val(gcLadeCom)
    frmWKL20!MSComm2.InputLen = 0
    frmWKL20!MSComm2.Settings = "9600,N,8,1"
    
    frmWKL20!MSComm2.RThreshold = 1
    If Not frmWKL20!MSComm2.PortOpen = True Then
        frmWKL20!MSComm2.PortOpen = True
    End If
    frmWKL20!MSComm2.Output = "AT" + Chr$(13)
    frmWKL20!MSComm2.PortOpen = False
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Then 'Anschluss bereits geöffnet
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul20"
        Fehler.gsFunktion = "OpenDrawerViaComPortModul20"
        Fehler.gsFehlertext = "Im Programteil Kasse/M20 ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If

End Sub
Public Function ermittleBonusjetzt() As Double
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim dWert       As Double
    Dim dKdBonus    As Double
    Dim dKJPreis    As Double
    
    ermittleBonusjetzt = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    dKdBonus = 0
    
    For lAktSatz = 0 To lAnzSatz - 1
        dKJPreis = 0
        cLBSatz = frmWKL20.List1.list(lAktSatz)

        
        '***************************************************
        '* Zeile ZWISCHENSUMME darf nicht übernommen werden!
        '***************************************************
        
        If ctmp <> "000000" Then
            
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
            dKJPreis = Val(ctmp)
            
            If Len(cLBSatz) >= 154 Then
                If Mid(cLBSatz, 154, 1) <> "N" Then
                    dKdBonus = dKdBonus + dKJPreis
                End If
                
            End If
        End If
    Next lAktSatz

    ermittleBonusjetzt = dKdBonus
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermittleBonusjetzt"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
    
End Function
Public Function sindAboPlusdrin() As Boolean
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim dAbo_Preis  As Double
    Dim lartnr      As Long
    Dim i           As Integer
    
    sindAboPlusdrin = False
    
    dAbo_Preis = 0
    gdABOPLUS_WERT = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        If ctmp <> "000000" Then
        
            ctmp = Mid(cLBSatz, 7, 6)
            lartnr = CLng(Trim$(ctmp))
            
            For i = 0 To 19
                If glWGTaste(i) = lartnr Then
                
                    ctmp = Mid(cLBSatz, 60, 9)
                    ctmp = Trim$(ctmp)
                    ctmp = fnMoveComma2Point$(ctmp)
                    dAbo_Preis = dAbo_Preis + Val(ctmp)
                    sindAboPlusdrin = True
                    Exit For
                End If
            Next i
            
        End If
    Next lAktSatz
    
    gdABOPLUS_WERT = dAbo_Preis

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "sindAboPlusdrin"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function sind_CouponBedingungen_erfüllt(iMindestMenge As Integer, dMindestUmsatz As Double, cCoupon_ID As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim ctmp            As String
    Dim cLBSatz         As String
    Dim i               As Integer
    Dim cEAN            As String
    Dim rsArt           As Recordset
    Dim dHenkelwert     As Double
    Dim iHenkelstück    As Integer
    Dim cSQL            As String
    
    dHenkelwert = 0
    iHenkelstück = 0
    
    sind_CouponBedingungen_erfüllt = False
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        If ctmp <> "000000" Then
        
            ctmp = Mid(cLBSatz, 7, 6)
            
            cEAN = ""
            
            loeschNEW "TEMP_EAN", gdBase
            cSQL = "Create table TEMP_EAN (Artnr long, EAN Text(13))"
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into TEMP_EAN Select artnr, ean from Artikel where Artnr = " & ctmp
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into TEMP_EAN Select artnr, ean2 as ean from Artikel where Artnr = " & ctmp
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into TEMP_EAN Select artnr, ean3 as ean from Artikel where Artnr = " & ctmp
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Insert into TEMP_EAN Select artnr, ean from Artean_k where Artnr = " & ctmp
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Delete * from TEMP_EAN  where ean is null "
            gdBase.Execute cSQL, dbFailOnError
            
            cSQL = "Delete * from TEMP_EAN  where val(ean) =0 "
            gdBase.Execute cSQL, dbFailOnError
            
            
            cSQL = "Select EAN from TEMP_EAN "
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
            
                rsArt.MoveFirst
                Do While Not rsArt.EOF
                
                    If Not IsNull(rsArt!EAN) Then
                        cEAN = rsArt!EAN
                        
                        If cEAN <> "" Then
            
                            cSQL = "select * from COUPONPRODUKTE where EAN = '" & cEAN & "'"
                            cSQL = cSQL & " and Coupon_ID = " & cCoupon_ID
                        
                            If DatendrinSQL(cSQL, gdBase) Then
                            
                                ctmp = Mid(cLBSatz, 60, 9)
                                ctmp = Trim$(ctmp)
                                ctmp = fnMoveComma2Point$(ctmp)
                                dHenkelwert = dHenkelwert + Val(ctmp)
                                
                                If Left(cLBSatz, 1) = "x" Then
                                    ctmp = Mid(cLBSatz, 2, 4)
                                Else
                                    ctmp = Mid(cLBSatz, 1, 5)
                                End If
                                ctmp = Trim$(ctmp)
                                iHenkelstück = iHenkelstück + Val(ctmp)
                                Exit Do
                        
                            End If
                        End If
                    End If
                
                rsArt.MoveNext
                Loop
            End If
            rsArt.Close: Set rsArt = Nothing
        End If
    Next lAktSatz
    
    If iHenkelstück >= iMindestMenge Then
        If dHenkelwert >= dMindestUmsatz Then
            sind_CouponBedingungen_erfüllt = True
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "sind_CouponBedingungen_erfüllt"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function Wie_oft_dieseCouponID_schon_drin(cBonusID As String) As Integer
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim ctmp            As String
    Dim cLBSatz         As String
    Dim sSQL As String
    
    Wie_oft_dieseCouponID_schon_drin = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        If ctmp <> "000000" Then
        
            ctmp = Mid(cLBSatz, 7, 6)
            
            sSQL = "Select Coupon_ID from COUPONREGELN where Coupon_artnr = " & ctmp
            
            If DatendrinSQL(sSQL, gdBase) Then
                Wie_oft_dieseCouponID_schon_drin = Wie_oft_dieseCouponID_schon_drin + 1
            Else
            
            End If
    
            
        End If
    Next lAktSatz

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Wie_oft_dieseCouponID_schon_drin"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function Wie_vielZehnProzLinr_schon_drin() As Integer
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim ctmp            As String
    Dim cSternchen      As String
    Dim cLBSatz         As String
    Dim lMenge          As Long
    Dim sSQL            As String
    
    Wie_vielZehnProzLinr_schon_drin = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        If ctmp <> "000000" Then
            ctmp = Mid(cLBSatz, 7, 6)
            
            sSQL = "Select artnr from Artlief where linr = " & glZehnProzLinr & " and artnr = " & ctmp
            
            If DatendrinSQL(sSQL, gdBase) Then
            
                'steht dort ein Sternchen?
                cSternchen = ""
                cSternchen = Mid(cLBSatz, 6, 1)
            
                'Menge
                lMenge = 0
            
                If Left(cLBSatz, 1) = "x" Then
                    ctmp = Mid(cLBSatz, 2, 4)
                Else
                    ctmp = Mid(cLBSatz, 1, 5)
                End If
                
                lMenge = CLng(ctmp)
                
                If Trim(cSternchen) = "" Then
                    Wie_vielZehnProzLinr_schon_drin = Wie_vielZehnProzLinr_schon_drin + lMenge
                End If
            Else
            
            End If
        End If
    Next lAktSatz

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Wie_vielZehnProzLinr_schon_drin"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function ZehnProzRechnen() As Double
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz            As Long
    Dim lAktSatz            As Long
    Dim ctmp                As String
    Dim cSternchen          As String
    Dim cLBSatz             As String
    Dim lMenge              As Long
    Dim sSQL                As String
    
    Dim dGesArtikelrabatt   As Double
    Dim dEinzelRabatt       As Double
    
    Dim dEinzelpreis        As Double
    Dim dGesamtPreis        As Double
    
    ZehnProzRechnen = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        If ctmp <> "000000" Then
        
            ctmp = Mid(cLBSatz, 7, 6)
'            MsgBox cLBSatz

            sSQL = "Select artnr from Artlief where linr = " & glZehnProzLinr & " and artnr = " & ctmp

            If DatendrinSQL(sSQL, gdBase) Then
            
                'steht dort ein Sternchen?
                cSternchen = ""
                cSternchen = Mid(cLBSatz, 6, 1)
            

                'Menge
                lMenge = 0
                
                dEinzelRabatt = 0
                dGesamtPreis = 0
                
                ctmp = Mid(cLBSatz, 50, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                dEinzelpreis = Val(ctmp)
                
                ctmp = Mid(cLBSatz, 60, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                dGesamtPreis = Val(ctmp)

                If Left(cLBSatz, 1) = "x" Then
                    ctmp = Mid(cLBSatz, 2, 4)
                Else
                    ctmp = Mid(cLBSatz, 1, 5)
                End If

                lMenge = CLng(ctmp)
                
                If Trim(cSternchen) = "" Then
                    'wenn kein Sternchen, dann rechnen wir den Rabatt aus
                    
                    dEinzelRabatt = dGesamtPreis * 10 / 100
                    dGesArtikelrabatt = dGesArtikelrabatt + dEinzelRabatt
                    
                    'und setzen ein Sternchen
                End If

            Else

            End If
    
            
        End If
    Next lAktSatz
    
    ZehnProzRechnen = dGesArtikelrabatt
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ZehnProzRechnen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function DieseArtikelImWarenkorb(cAusnahmeArtikel As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim ctmp            As String
    Dim cLBSatz         As String
    Dim i               As Integer
    Dim cFeld           As String
    
    DieseArtikelImWarenkorb = False
    
    If cAusnahmeArtikel = "" And gbNurBonusfRunden = False Then
        Exit Function
    End If
    
    
'    gbNurBonusfRunden






    If gbNurBonusfRunden = True Then
        lAnzSatz = frmWKL20.List1.ListCount
        
        For lAktSatz = 0 To lAnzSatz - 1
            
            cLBSatz = frmWKL20.List1.list(lAktSatz)
            If ctmp <> "000000" Then
            
                If Mid(cLBSatz, 6, 1) = "*" Then
                    DieseArtikelImWarenkorb = True
                    Exit For
                End If
            End If
        
        Next lAktSatz
    End If
    
    If DieseArtikelImWarenkorb = True Then
        Exit Function
    End If


    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        If ctmp <> "000000" Then
            
            ctmp = Trim(Mid(cLBSatz, 7, 6))
            
            Dim sArray() As String
            sArray = Split(cAusnahmeArtikel, "$")
            
            For i = 0 To UBound(sArray)
                cFeld = sArray(i)
                If ctmp = cFeld Then
                    DieseArtikelImWarenkorb = True
                    Exit For
                End If
            Next i
        End If
    
        If DieseArtikelImWarenkorb = True Then
            Exit For
        End If
    Next lAktSatz

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "DieseArtikelImWarenkorb"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function ermittleVKMenge_bestimmte_Artnr_drin(lSuch_Artnr As Long) As Long
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim lartnr      As Long
    Dim lMenge      As Long
    
    ermittleVKMenge_bestimmte_Artnr_drin = 0
    
    lMenge = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        
        ctmp = Mid(cLBSatz, 7, 6)
        lartnr = CLng(Trim$(ctmp))
        
        If lSuch_Artnr = lartnr Then
            'summiere die Menge
            
            
            
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If
            
            lMenge = CLng(ctmp)
            ermittleVKMenge_bestimmte_Artnr_drin = ermittleVKMenge_bestimmte_Artnr_drin + lMenge
            
        End If
    Next lAktSatz
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermittleVKMenge_bestimmte_Artnr_drin"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function ermittleVKMenge_UeberStaffelNR(lStaffelnr As Long) As Long
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim lartnr      As Long
    Dim lMenge      As Long
    
    ermittleVKMenge_UeberStaffelNR = 0
    
    lMenge = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        
        ctmp = Mid(cLBSatz, 7, 6)
        lartnr = CLng(Trim$(ctmp))
        
        If Gehoerst_du_zu_Staffel(lartnr, lStaffelnr) Then
            'summiere die Menge
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If
            
            lMenge = CLng(ctmp)
            ermittleVKMenge_UeberStaffelNR = ermittleVKMenge_UeberStaffelNR + lMenge
            
        End If
    Next lAktSatz
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermittleVKMenge_UeberStaffelNR"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function Gehoerst_du_zu_Staffel(lArt As Long, lStaffNr As Long) As Boolean
On Error GoTo LOKAL_ERROR

    Gehoerst_du_zu_Staffel = False
    
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    sSQL = "Select * from STAFFEL_KVK_ARTIKEL where STAFFNR = " & lStaffNr & " and Artnr = " & lArt
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Gehoerst_du_zu_Staffel = True
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Gehoerst_du_zu_Staffel"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function sind_bestimmte_Artnr_drin(lSuch_Artnr As Long) As Boolean
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim lartnr      As Long
    
    sind_bestimmte_Artnr_drin = False
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        
        ctmp = Mid(cLBSatz, 7, 6)
        lartnr = CLng(Trim$(ctmp))
        
        If lSuch_Artnr = lartnr Then
            sind_bestimmte_Artnr_drin = True
            Exit For
        End If
    Next lAktSatz
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "sind_bestimmte_Artnr_drin"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function nurStornofArtikel_enthalten() As Boolean
On Error GoTo LOKAL_ERROR

    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim lartnr      As Long
    
    Dim lMenge      As Long
    Dim rsrs        As DAO.Recordset
    Dim sSQL        As String
    
    nurStornofArtikel_enthalten = True
    
    lAnzSatz = frmWKL20.List1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        
        ctmp = Mid(cLBSatz, 7, 6)
        lartnr = CLng(Trim$(ctmp))
        
        If Left(cLBSatz, 1) = "x" Then
            ctmp = Mid(cLBSatz, 2, 4)
        Else
            ctmp = Mid(cLBSatz, 1, 5)
        End If

        ctmp = Trim$(ctmp)
        ctmp = fnMoveComma2Point$(ctmp)
        lMenge = Val(ctmp)
        
        If lMenge < 0 Then

            ctmp = "J"
        
            sSQL = "Select MERK from STORNOF where ARTNR = " & lartnr
            Set rsrs = gdBase.OpenRecordset(sSQL)
    
            If Not rsrs.EOF Then
       
                If Not IsNull(rsrs!merk) Then
                    ctmp = rsrs!merk
                End If
                
                If ctmp = "N" Then
                    nurStornofArtikel_enthalten = False
                    rsrs.Close: Set rsrs = Nothing
                    Exit Function
                End If
            End If
    
            rsrs.Close: Set rsrs = Nothing
        
        End If
        
        
        
        
        
        
        
        
    Next lAktSatz
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "nurStornofArtikel_enthalten"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Function ermfilfromKUNDE(gsKU As String) As Integer
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermfilfromKUNDE = -1
    
    If Trim(gsKU) = "" Then
        Exit Function
    End If
    
    cSQL = "Select * from kunden where kundnr = " & gsKU
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!FILIALNR) Then
            ermfilfromKUNDE = rsrs!FILIALNR
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermfilfromKUNDE"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermittleBonusVorher(gsKU As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    If Trim(gsKU) = "" Then
        Exit Function
    End If
    
    ermittleBonusVorher = 0
    
    cSQL = "Select * from kunden where kundnr = " & gsKU
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!BONUS) Then
            ermittleBonusVorher = rsrs!BONUS
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermittleBonusVorher"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function KundeBonusfähig(sKUNDNR As String) As Boolean
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cMerkmal2 As String
    
    If sKUNDNR = "" Then
        Exit Function
    End If
    
    KundeBonusfähig = True
    cMerkmal2 = "J"
    
    cSQL = "Select Merkmal2 from kunden where kundnr = " & sKUNDNR
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!MERKMAL2) Then
            cMerkmal2 = rsrs!MERKMAL2
            
            If cMerkmal2 = "N" Then
                KundeBonusfähig = False
            End If
            
        Else
            KundeBonusfähig = True
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "KundeBonusfähig"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function WasistmitBonus() As Boolean
On Error GoTo LOKAL_ERROR

WasistmitBonus = False

gdbonusGutschein = 0



Dim dBonusvomKundevoher As Double
Dim dBonusfaehigjetzt As Double

dBonusvomKundevoher = 0
dBonusfaehigjetzt = 0

If gbbonusHerab = False Then
    If gbBonusBNB = False Then 'Bonusüberprüfung sofort
        If frmWKL20.Label2(7).Caption <> "0" Then 'ist ein Kunde gewählt
            gckundnr = Trim(frmWKL20.Label2(7).Caption)
            If gbFILBONI = True Then
                If ermfilfromKUNDE(frmWKL20.Label2(7).Caption) = CInt(gcFilNr) Then
                
                    dBonusfaehigjetzt = ermittleBonusjetzt
                    dBonusvomKundevoher = ermittleBonusVorher(frmWKL20.Label2(7).Caption)
                    
                    dBonusfaehig = dBonusfaehigjetzt + dBonusvomKundevoher
                    
                    If gdBonusGrenze > 0 Then
                        If dBonusfaehig >= gdBonusGrenze Then
                            If KundeBonusfähig(frmWKL20.Label2(7).Caption) Then
                                frmWK20h.Show 1
                            End If
                            
                            If gbbonusHerab = True Then
                                WasistmitBonus = False
                                Exit Function
                            End If
                        End If
                    End If
                    
                    
                End If
            Else
            
                dBonusfaehigjetzt = ermittleBonusjetzt
                dBonusvomKundevoher = ermittleBonusVorher(frmWKL20.Label2(7).Caption)
                
                dBonusfaehig = dBonusfaehigjetzt + dBonusvomKundevoher
                
                If gdBonusGrenze > 0 Then
                    If dBonusfaehig >= gdBonusGrenze Then
                
    '                    MsgBox dBonusfaehig
                        If KundeBonusfähig(frmWKL20.Label2(7).Caption) Then
                            frmWK20h.Show 1
                        End If
                        
                        If gbbonusHerab = True Then
                            WasistmitBonus = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
End If

WasistmitBonus = True

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "WasistmitBonus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Function
Public Sub Bonus_neu_AA()
On Error GoTo LOKAL_ERROR

Dim dBonusvomKundevoher As Double
Dim dBonusfaehigjetzt As Double

giAnzBonus_Erreicht = 0

dBonusvomKundevoher = 0
dBonusfaehigjetzt = 0

If gbBonusBNB = False Then 'Bonusüberprüfung sofort
    If frmWKL20.Label2(7).Caption <> "0" Then 'ist ein Kunde gewählt
        gckundnr = Trim(frmWKL20.Label2(7).Caption)
        If gbFILBONI = True Then
            If ermfilfromKUNDE(frmWKL20.Label2(7).Caption) = CInt(gcFilNr) Then
            
                dBonusfaehigjetzt = ermittleBonusjetzt
                dBonusvomKundevoher = ermittleBonusVorher(frmWKL20.Label2(7).Caption)
                
                dBonusfaehig = dBonusfaehigjetzt + dBonusvomKundevoher
                
                If dBonusfaehig >= gdBonusGrenze Then
                    
                    If KundeBonusfähig(frmWKL20.Label2(7).Caption) Then
                        frmWK20i.Show 1
                    End If
                End If
            End If
        Else
            dBonusfaehigjetzt = ermittleBonusjetzt
            dBonusvomKundevoher = ermittleBonusVorher(frmWKL20.Label2(7).Caption)
            
            dBonusfaehig = dBonusfaehigjetzt + dBonusvomKundevoher
            
            If dBonusfaehig >= gdBonusGrenze Then

                If KundeBonusfähig(frmWKL20.Label2(7).Caption) Then
                    frmWK20i.Show 1
                End If
                
            End If
        End If
    End If
End If

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Bonus_neu_AA"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Sub
Public Sub WasistmitPLZ_Erfragen(sKUNDNR As String, dMoney As Double)
On Error GoTo LOKAL_ERROR

    Dim sPlz As String
    Dim sSQL As String
    Dim cDatum As String
    cDatum = Fix(Now)
    
    Dim czeit As String
    czeit = Format$(Now, "HH:MM:SS")
    
   
    If gbPLZGEBIET_AuchBeiKUWAHL = True Then
    
    Else
        If Val(sKUNDNR) > 0 Then
            Exit Sub
        End If
    End If
    
    gsPLZ = ""
    frmWKL176.Show 1
    
   
    
    If Val(gsPLZ) > 0 Then
        sSQL = "Insert into PLZGEBIET (PLZ,ADATE,AZEIT,FILIALE,GELDWERT,SENDOK) "
        sSQL = sSQL & " values ('" & gsPLZ & "','" & cDatum & "', '" & czeit & "'," & CByte(gcFilNr) & ",'" & dMoney & "',False)"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "WasistmitPLZ_Erfragen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Sub

Public Sub WasistmitAboPlus()
On Error GoTo LOKAL_ERROR

    Dim i                   As Integer
    Dim rsrs                As Recordset
    Dim sSQL                As String

    If NewTableSuchenDBKombi("ABOPLUS_UMS", gdBase) = False Then
        CreateTableT2 "ABOPLUS_UMS", gdBase
    End If
    
    For i = 0 To 19
        glWGTaste(i) = 0
    Next i
    
    Set rsrs = gdBase.OpenRecordset("ABOPLUS")
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        i = 0
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                glWGTaste(i) = rsrs!artnr
                i = i + 1
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    If sindAboPlusdrin Then
    
        'scann die Karte
        frmWKL175.Show 1
        
        'und eintragen
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "WasistmitAboPlus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Sub
'Public Sub WasistmitGutscheinAngebot()
'On Error GoTo LOKAL_ERROR
'
'        dlgGutschein.Show 1
'
'        Select Case dlgGutschein.Back
'            Case 1 'Ja
''                    HoleUnterbrochenenBonWK20b_alsGS svorgangsnummer
'            Case 2 'Nein
''                    HoleUnterbrochenenBonWK20b svorgangsnummer
'        End Select
'
'Exit Sub
'LOKAL_ERROR:
'
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = "Modul20"
'    Fehler.gsFunktion = "WasistmitGutscheinAngebot"
'    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
'
'    Fehlermeldung1
'End Sub
Public Sub WasistmitBONUSNRE()
On Error GoTo LOKAL_ERROR

Dim rsrs                As Recordset
Dim sSQL                As String
Dim lBONUSNRE_Artnr     As Long
    
If NewTableSuchenDBKombi("BONUSNR", gdBase) = False Then
    CreateTable "BONUSNR", gdBase
End If

Set rsrs = gdBase.OpenRecordset("BONUSNRE")
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!artnr) Then
        lBONUSNRE_Artnr = rsrs!artnr
    End If
End If
rsrs.Close

If sind_bestimmte_Artnr_drin(lBONUSNRE_Artnr) Then
    'hier die Nummer vom Bonusschreiben eingeben
    frmWKL177.Show 1
    'hier die Nummer vom Bonusschreiben eingeben
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "WasistmitBONUSNRE"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Sub

Public Sub WasistmitBONUSAA()
On Error GoTo LOKAL_ERROR

Dim rsrs    As Recordset
glBONUSAA_Artnr = 0

If NewTableSuchenDBKombi("BONUSAA", gdBase) = True Then
    Set rsrs = gdBase.OpenRecordset("BONUSAA")
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        If Not IsNull(rsrs!artnr) Then
            glBONUSAA_Artnr = Val(rsrs!artnr)
        End If
    End If
    rsrs.Close
End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "WasistmitBONUSAA"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Sub
Public Sub ist_Aboplus()
On Error GoTo LOKAL_ERROR

Dim rsrs    As Recordset
gbABOPLUS = False

If NewTableSuchenDBKombi("ABOPLUS", gdBase) = True Then
    gbABOPLUS = True
End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ist_Aboplus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Sub
Public Sub ist_Bonus_nummer()
On Error GoTo LOKAL_ERROR

Dim rsrs    As Recordset
glBONUSNRE_Artnr = 0

If NewTableSuchenDBKombi("BONUSNRE", gdBase) = True Then
    Set rsrs = gdBase.OpenRecordset("BONUSNRE")
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        If Not IsNull(rsrs!artnr) Then
            glBONUSNRE_Artnr = Val(rsrs!artnr)
        End If
    End If
    rsrs.Close
End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ist_Bonus_nummer"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasse auf."
    
    Fehlermeldung1
End Sub
Public Function ermUste(sTab As String) As Double
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim ess     As Single
    Dim ess100  As Single
    Dim R       As Recordset
    
    ess = gdMWStE
    ess100 = ess + 100

    ermUste = 0

    sSQL = "select (sum(vkpr * anzahl) * " & ess & ") /" & ess100 & " as erm from " & sTab & " where mwst = 'E' "

    FnOpenrecordset R, sSQL, 1, gdBase
    If Not R.EOF Then
        If Not IsNull(R!ERM) Then
            ermUste = R!ERM
        End If
    End If
    R.Close
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermUste"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermUstv(sTab As String) As Double
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim vss As Single
Dim vss100 As Single
Dim R As Recordset

vss = gdMWStV
vss100 = vss + 100

ermUstv = 0

sSQL = "select (sum(vkpr * anzahl) * " & vss & ") /" & vss100 & " as voll from " & sTab & " where mwst = 'V' "
FnOpenrecordset R, sSQL, 1, gdBase

If Not R.EOF Then
    If Not IsNull(R!VOLL) Then
        ermUstv = R!VOLL
    End If
End If
R.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermUstv"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub lese_Storno_Text_in_Array()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As DAO.Recordset
    
    If NewTableSuchenDBKombi("STORNOTEXT", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("Select * from STORNOTEXT order by ZNR")
        If Not rsrs.EOF Then
            rsrs.MoveLast
            
            ReDim sStornoText(rsrs.RecordCount)
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!ZTEXT) Then
                    sStornoText(rsrs!ZNR) = rsrs!ZTEXT
                End If
            rsrs.MoveNext
            Loop
        Else
            ReDim sStornoText(0)
            sStornoText(0) = ""
        End If
        rsrs.Close: Set rsrs = Nothing
    Else
        ReDim sStornoText(3)
        sStornoText(0) = "Ansonsten müssen wir Ihnen"
        sStornoText(1) = "70 % des Behandlungspreises ver-"
        sStornoText(2) = "rechnen. Wir danken für Ihr Ver-"
        sStornoText(3) = "ständnis!"
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "lese_Storno_Text_in_Array"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub lese_Storno_Text_in_Array_T1()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As DAO.Recordset
    
    If NewTableSuchenDBKombi("STORNOTEXTT1", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("Select * from STORNOTEXTT1 order by ZNR")
        If Not rsrs.EOF Then
            rsrs.MoveLast
            
            ReDim sStornoTextT1(rsrs.RecordCount)
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!ZTEXT) Then
                    sStornoTextT1(rsrs!ZNR) = rsrs!ZTEXT
                End If
            rsrs.MoveNext
            Loop
            
        End If
        rsrs.Close: Set rsrs = Nothing
    Else
        ReDim sStornoTextT1(7)
        
        sStornoTextT1(0) = "Bitte kommen Sie rechtzeitig vor"
        sStornoTextT1(1) = "Ihrem Behandlungstermin. Ver-"
        sStornoTextT1(2) = "spätungen haben leider eine"
        sStornoTextT1(3) = "kürzere Behandlung zur Folge."
        sStornoTextT1(4) = "Wenn Sie einen Termin nicht"
        sStornoTextT1(5) = "einhalten können, bitten wir Sie"
        sStornoTextT1(6) = "mindestens 3 Tage davor "
        sStornoTextT1(7) = "abzusagen."
    
    End If
    
        
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "lese_Storno_Text_in_Array_T1"
    Fehler.gsFehlertext = "Im Programmteil Tabellenbearbeitung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub DruckeZweitBonAusListe(Listx As ListBox, bMitLeerZeilen)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz            As Long
    Dim lAktSatz            As Long
    Dim lcount              As Long
    Dim lAnzZeile           As Long
    Dim cLBSatz             As String
    Dim cDaten              As String
    Dim aDeviceName         As String
    Dim cEscapeSequenz      As String
    Dim iLenZeile           As Integer
    Dim dDruckRabattWert    As Double
    

''''''    setzedrucker gcBonDrucker

    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz
    DoEvents
    
    dDruckRabattWert = 0
    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    cDaten = ""
    iLenZeile = 32

    'Listentext an Drucker senden
    lAnzSatz = Listx.ListCount
    
    
    
   
    
    
    'Bild
    If gcBild <> "" Then
'        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    'ende bild
    

    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    dDruckRabattWert = ermittleRabattwert(Listx)
    
    For lAktSatz = 0 To lAnzSatz - 1
        If dDruckRabattWert > 0 And lAktSatz > lAnzSatz - 10 Then
            Exit For
        End If
        cLBSatz = Listx.list(lAktSatz)
        
        cDaten = cLBSatz
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lAktSatz
    
    If bMitLeerZeilen Then
        For lcount = 1 To 9
            If lcount = 9 Then
                cEscapeSequenz = "." & vbCrLf
            Else
                cEscapeSequenz = " " & vbCrLf
            End If
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        Next lcount
    End If
    
    ''''''''''''''''''''''' Oday Neu ''''''''''''''''''''''''''
    '******************************************************
    'check ob Bon TSE hat  <<<<< Teil1 <<<<<<< START
    '******************************************************
        Dim HatTSE As Boolean
        Dim j As Integer
        
        For j = 1 To UBound(cDruckZeile)
          If InStr(1, cDruckZeile(j), "TSE Start: ") Then
           HatTSE = True
          End If
        Next j
    '*****************************************************
    'check ob Bon TSE hat <<<<<Teil1 <<<<<<< ENDE
    '*****************************************************
    ''''''''''''''''''''''' Oday Neu ''''''''''''''''''''''''''
    
'BON_DRUCKEN:
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        DoEvents
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        DoEvents
    End If
    
'eventuellen BonusBarcode drucken und an die Leerzeilen denken

    If dDruckRabattWert > 0 Then
        Barcode_Bonus CStr(dDruckRabattWert), "7"
    End If
     
    
    'Druckbereich freigeben
    Erase cDruckZeile
    DoEvents
    
    
    ''''''''''''''''''''''' Oday Neu '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If MitQrCode And E_TSE_Aktiv And gbQRFlag Then

        'beim altDruckModus kann der Drucker beim Kunde kein QR-Code drucken (alte Drucker)
        If altDruckModus Then

        Else
            QRcodeDrucken
            Sleep 1000
        End If
 
    End If
 
    
    '******************************************************
    'ein paar Leerzeilen drucken  <<<<< Teil1 <<<<<<< START
    '******************************************************
    If Not bMitLeerZeilen And HatTSE Then
        lAnzZeile = 0
        For lcount = 1 To 9
            If lcount = 9 Then
               cEscapeSequenz = vbCrLf
            Else
               cEscapeSequenz = " " & vbCrLf
            End If
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        Next lcount
        
            'BON_DRUCKEN:
            If gbAPI = True Then
                OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
                DoEvents
            Else
                OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
                DoEvents
            End If
    End If
    '******************************************************
    'ein paar Leerzeilen drucken  <<<<< Teil1 <<<<<<< ENDE
    '******************************************************
    
    ''''''''''''''''''''''' Oday Neu '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
       
    
 'BON_SCHNEIDEN:
    If altDruckModus Then
       'Papier schneiden (alte Funktion)
          If gbAPI = True Then
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcSchneiden
            OpenDrawer aDeviceName, cEscapeSequenz
          End If
    Else
        'Papier schneiden (neue Funktion)
         CutPapier
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "DruckeZweitBonAusListe"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function Hole_Volle_MWST_aus_Bon(Listx As ListBox) As Double
    On Error GoTo LOKAL_ERROR
    
    Hole_Volle_MWST_aus_Bon = 0
    
    Dim iCount As Long
    Dim cZeile As String
    For iCount = 0 To Listx.ListCount - 1
        cZeile = Listx.list(iCount)
        
        If Left(cZeile, 21) = "MWSt.-Anteil: " & gdMWStV & "% EUR" Then
            Hole_Volle_MWST_aus_Bon = Trim(Right(cZeile, Len(cZeile) - 21))
            Exit For
        ElseIf Left(cZeile, 17) = "MWSt.-Anteil: " & gdMWStV & "%" Then
            
            Hole_Volle_MWST_aus_Bon = Trim(Right(cZeile, Len(cZeile) - 17))
            Exit For
        End If
    Next iCount
    
    
   
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Hole_Volle_MWST_aus_Bon"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Hole_Erm_MWST_aus_Bon(Listx As ListBox) As Double
    On Error GoTo LOKAL_ERROR
    
    Hole_Erm_MWST_aus_Bon = 0
    
    Dim iCount As Long
    Dim cZeile As String
    For iCount = 0 To Listx.ListCount - 1
        cZeile = Listx.list(iCount)
        
        If Left(cZeile, 21) = "MWSt.-Anteil: " & gdMWStE & "%  EUR" Then
            Hole_Erm_MWST_aus_Bon = Trim(Right(cZeile, Len(cZeile) - 21))
            Exit For
        ElseIf Left(cZeile, 16) = "MWSt.-Anteil: " & gdMWStE & "%" Then
            Hole_Erm_MWST_aus_Bon = Trim(Right(cZeile, Len(cZeile) - 16))
            Exit For
        End If
    Next iCount
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Hole_Erm_MWST_aus_Bon"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermittleRabattwert(Listx As ListBox) As Double
On Error GoTo LOKAL_ERROR

    ermittleRabattwert = 0
    
    If gsWWBonusArtnr = "0" Then
        Exit Function
    End If

    Dim lPos                As Long
    Dim cZeichen            As String
    Dim cRabattwert         As String
    Dim i                   As Integer
    Dim iCount              As Integer
    Dim cZeile              As String
    Dim iZaehler            As Integer
    
    For iCount = 0 To Listx.ListCount - 1
        cZeile = Listx.list(iCount)
        
        If cZeile = "---------- Gutschein -----------" Then
            'Ja Gutschein ist gegeben
            iZaehler = iCount
            Exit For
        End If
    Next iCount
    
    If iZaehler = 0 Then
        Exit Function
    End If
    
    For iCount = iZaehler To Listx.ListCount - 1
        cZeile = Listx.list(iCount)
        
        If InStr(1, cZeile, ",") Then 'ist in der Zeile ein Komma dann
        
            cRabattwert = ""
            lPos = InStr(1, cZeile, ",") 'position vom Komma
            
            
            If lPos > 2 Then
                lPos = lPos - 3
            
                If lPos = 0 Then
                
                    lPos = 1
                    For i = lPos To lPos + 4
                        cZeichen = Mid(cZeile, i, 1)
                        
                        If InStr(" ", cZeichen) > 0 Then
                            cRabattwert = ""
                        End If
                        
                        If InStr("1234567890,", cZeichen) > 0 Then
                            cRabattwert = cRabattwert & cZeichen
                        End If
                    Next i
                
                Else
            
                    For i = lPos To lPos + 5
                        cZeichen = Mid(cZeile, i, 1)
                        
                        If InStr(" ", cZeichen) > 0 Then
                            cRabattwert = ""
                        End If
                        
                        If InStr("1234567890,", cZeichen) > 0 Then
                            cRabattwert = cRabattwert & cZeichen
                        End If
                    Next i
                End If
            
            
            
            ElseIf lPos > 1 Then
                lPos = lPos - 2
                
                If lPos = 0 Then
                
                    lPos = 1
                    For i = lPos To lPos + 3
                        cZeichen = Mid(cZeile, i, 1)
                        
                        If InStr(" ", cZeichen) > 0 Then
                            cRabattwert = ""
                        End If
                        
                        If InStr("1234567890,", cZeichen) > 0 Then
                            cRabattwert = cRabattwert & cZeichen
                        End If
                    Next i
                
                Else
            
                    For i = lPos To lPos + 4
                        cZeichen = Mid(cZeile, i, 1)
                        
                        If InStr(" ", cZeichen) > 0 Then
                            cRabattwert = ""
                        End If
                        
                        If InStr("1234567890,", cZeichen) > 0 Then
                            cRabattwert = cRabattwert & cZeichen
                        End If
                    Next i
                End If
            
            End If
            
            
            
            
            
            
            
            
            If IsNumeric(cRabattwert) Then
                ermittleRabattwert = CDbl(cRabattwert)
                Exit For
            End If
        End If
    Next iCount

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermittleRabattwert"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Function
Public Sub Barcode_Bonus(cSumme As String, cVorgang As String)
On Error GoTo LOKAL_ERROR

    Dim lcount                  As Long
    Dim cZeichen                As String
    Dim cArtNr                  As String
    Dim aDeviceName             As String
    Dim cEscapeSequenz          As String
    Dim lPruefZiffer            As Long

    aDeviceName = Printer.DeviceName
    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(27) & Chr(97) & Chr(1)
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = Chr(29) & Chr(72) & Chr(2)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    'die Barcodehöhe
    
'    cEscapeSequenz = Chr(29) & Chr(104) & Chr(164)
'    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(104) & Chr(40)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = Chr(29) & Chr(119) & Chr(3)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    Dim cFilnrDruck As String
    
    If Len(gcFilNr) = 1 Then
        cFilnrDruck = "0" & gcFilNr
    ElseIf Len(gcFilNr) = 2 Then
        cFilnrDruck = gcFilNr
    End If
    
    Dim lBonus_NR As Long
    lBonus_NR = insert_BONUS_SYS_Back_Max(CDbl(cSumme))
    
    Dim cBonus_NR As String
    cBonus_NR = CStr(lBonus_NR)
    
    If Len(cBonus_NR) = 1 Then
        cBonus_NR = "00000" & cBonus_NR
    ElseIf Len(cBonus_NR) = 2 Then
        cBonus_NR = "0000" & cBonus_NR
    ElseIf Len(cBonus_NR) = 3 Then
        cBonus_NR = "000" & cBonus_NR
    ElseIf Len(cBonus_NR) = 4 Then
        cBonus_NR = "00" & cBonus_NR
    ElseIf Len(cBonus_NR) = 5 Then
        cBonus_NR = "0" & cBonus_NR
    ElseIf Len(cBonus_NR) = 6 Then

    End If
    
    cArtNr = "2" & "50" & cFilnrDruck & cVorgang & cBonus_NR
    
    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim p5 As Integer
    Dim p6 As Integer
    Dim p7 As Integer
    Dim p8 As Integer
    Dim p9 As Integer
    Dim p10 As Integer
    Dim p11 As Integer
    Dim p12 As Integer
    Dim p13 As Integer
    
    Dim rest As Double
    Dim pz As Long
    
    p1 = Val(Mid(cArtNr, 1, 1)) * 1
    p2 = Val(Mid(cArtNr, 2, 1)) * 3
    p3 = Val(Mid(cArtNr, 3, 1)) * 1
    p4 = Val(Mid(cArtNr, 4, 1)) * 3
    p5 = Val(Mid(cArtNr, 5, 1)) * 1
    p6 = Val(Mid(cArtNr, 6, 1)) * 3
    p7 = Val(Mid(cArtNr, 7, 1)) * 1
    p8 = Val(Mid(cArtNr, 8, 1)) * 3
    p9 = Val(Mid(cArtNr, 9, 1)) * 1
    p10 = Val(Mid(cArtNr, 10, 1)) * 3
    p11 = Val(Mid(cArtNr, 11, 1)) * 1
    p12 = Val(Mid(cArtNr, 12, 1)) * 3
    p13 = p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8 + p9 + p10 + p11 + p12
    
    rest = p13 Mod 10
    pz = 10 - rest
    If rest = 0 Then
        pz = 0
    End If

    cArtNr = cArtNr & Trim$(Str$(pz))
    
    cEscapeSequenz = gcBarCode & cArtNr
    OpenDrawer aDeviceName, cEscapeSequenz
    
    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz

    cEscapeSequenz = vbCrLf
    OpenDrawer aDeviceName, cEscapeSequenz
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "Barcode_Bonus"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Function seeklinr(searchstring As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    seeklinr = searchstring
    
    cSQL = "Select Linr,LIEFBEZ from LISRT where Kuerzel like  '*" & searchstring & "*'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount = 1 Then
            If Not IsNull(rsrs!linr) Then
                seeklinr = rsrs!linr
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    If searchstring <> seeklinr Then
        Exit Function
    End If
    
    cSQL = "Select Linr,LIEFBEZ from LISRT where LIEFBEZ like  '*" & searchstring & "*'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount = 1 Then
        
            If Not IsNull(rsrs!linr) Then
                seeklinr = rsrs!linr
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "seeklinr"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function seekliefbez(searchstring As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    seekliefbez = searchstring
    
    cSQL = "Select Linr,LIEFBEZ from LISRT where Kuerzel like  '*" & searchstring & "*'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount = 1 Then
        
            If Not IsNull(rsrs!LIEFBEZ) Then
                seekliefbez = rsrs!LIEFBEZ
            End If
            
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    If searchstring <> seekliefbez Then
        Exit Function
    End If
    
    cSQL = "Select Linr,LIEFBEZ from LISRT where LIEFBEZ like  '*" & searchstring & "*'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount = 1 Then
            If Not IsNull(rsrs!LIEFBEZ) Then
                seekliefbez = rsrs!LIEFBEZ
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "seekliefbez"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function seekPGNnr(searchstring As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    seekPGNnr = searchstring
    
    cSQL = "Select PGN from PGNDBF where PGNBEZEICH like  '*" & searchstring & "*'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount = 1 Then
        
            If Not IsNull(rsrs!PGN) Then
                seekPGNnr = rsrs!PGN
            End If
            
        End If
    End If
    rsrs.Close: Set rsrs = Nothing


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "seekPGNnr"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function seekPGNbez(searchstring As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    seekPGNbez = searchstring
    
    
    cSQL = "Select PGNBEZEICH from PGNDBF where PGNBEZEICH like  '*" & searchstring & "*'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount = 1 Then
            If Not IsNull(rsrs!PGNBEZEICH) Then
                seekPGNbez = rsrs!PGNBEZEICH
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "seekPGNbez"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub LiefKuerzelAufloesung(labelx As Label, textx As TextBox)
On Error GoTo LOKAL_ERROR

    Dim searchstr As String
    Dim sNeuLinr As String

    If Len(textx.Text) = 0 Then
        labelx.Caption = "kein Lieferant"
        labelx.Refresh
    End If

    If IsNumeric(textx.Text) = False Then
        'jetzt nach eindeutigen Lieferantennummern suchen
        If Len(textx.Text) > 0 And Len(textx.Text) < 10 Then

            searchstr = textx.Text
            sNeuLinr = seeklinr(searchstr)

            If IsNumeric(sNeuLinr) Then
                If sNeuLinr <> searchstr Then
                    textx.Text = sNeuLinr

                    labelx.Caption = seekliefbez(searchstr)
                    labelx.Refresh
                    anzeige "laser", "", frmWKL00.lbl6(28)

                End If
            Else
                labelx.Caption = "kein Lieferant"
                labelx.Refresh
            End If
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "LiefKuerzelAufloesung"
    Fehler.gsFehlertext = "Im Programmteil Lieferantenauflösung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub InsertAnsKassBuch(cKas As String)
On Error GoTo LOKAL_ERROR

Dim sSQL                    As String
Dim rsrs                    As Recordset
Dim lDate                   As Long
Dim iPos                    As Integer
Dim dWECHSEL                As Double
Dim dABSCHOPF               As Double
Dim dBargeld                As Double
Dim dBargeldKredittilgung   As Double
Dim dGutschAuszahlbetrag    As Double
Dim dKarte                  As Double
Dim dGutschVK               As Double
Dim dGutsch                 As Double
Dim dDUKA                   As Double
Dim dScheck                 As Double
Dim dLast                   As Double
Dim dKredit                 As Double
Dim dVKnichtUmsatz          As Double
Dim i                       As Integer

Screen.MousePointer = 11

lDate = DateValue(Now)

sSQL = " Select sum(UMS_BAR) as umsatzbar "
sSQL = sSQL & " ,sum(UMS_KRED) as umsatzkredit "
sSQL = sSQL & " ,sum(UMS_Karte) as umsatzkarte "
sSQL = sSQL & " ,sum(UMS_last) as umsatzlast "
sSQL = sSQL & " ,sum(UMS_scheck) as umsatzscheck "

sSQL = sSQL & " ,SUM(GUTSCHBAR) as vkGutschbar "
sSQL = sSQL & " ,sum(TILGBAR) as sTILGBAR "
sSQL = sSQL & " ,sum(AUSZGUTSCH) as SAUSZGUTSCH "


sSQL = sSQL & " ,sum(abschopf) as sumabschopf "
sSQL = sSQL & " ,SUM(DUKA) as umsatzDUKATEN "
sSQL = sSQL & " ,sum(wechsel) as sumwechsel "
sSQL = sSQL & " ,sum(ZHLGGUTSCH) as umsatzgutschein "
sSQL = sSQL & " from afcstat where Kasnum = " & cKas
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    
    iPos = 0
    dSumBar = 0
    dSumUms = 0
    
    'Wechselgeld
    dWECHSEL = 0
    dWECHSEL = ermLastWechselbetrag(CByte(gcKasNum))
    
    dSumBar = dSumBar + dWECHSEL
    
    iPos = iPos + 1
    InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Wechselgeld", 0, "", dWECHSEL, 0
    
    'Bargeld
    dBargeld = 0
    If Not IsNull(rsrs!umsatzbar) Then
        dBargeld = rsrs!umsatzbar
    Else
        dBargeld = 0
    End If
    
    dSumUms = dSumUms + dBargeld
    dSumBar = dSumBar + dBargeld
    
    If dBargeld <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Bar", dBargeld, "", dBargeld, 0
    End If
    
    'Tilgung in Bar
    dBargeldKredittilgung = 0
    If Not IsNull(rsrs!STILGBAR) Then
        dBargeldKredittilgung = rsrs!STILGBAR
    Else
        dBargeldKredittilgung = 0
    End If
    
    dSumBar = dSumBar + dBargeldKredittilgung
    
    If dBargeldKredittilgung <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Kredittilgung", 0, "", dBargeldKredittilgung, 0
    End If
    
    'neu Gutschein Auszahlung in Bar 05.12.2014
    dGutschAuszahlbetrag = 0
    If Not IsNull(rsrs!SAUSZGUTSCH) Then
        dGutschAuszahlbetrag = rsrs!SAUSZGUTSCH
    Else
        dGutschAuszahlbetrag = 0
    End If
    
    dGutschAuszahlbetrag = dGutschAuszahlbetrag * -1
    
    dSumBar = dSumBar + dGutschAuszahlbetrag
    
    If dGutschAuszahlbetrag <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Gutschein Auszahlungen", 0, "", dGutschAuszahlbetrag, 0
    End If
    
    'Ende neu Gutschein Auszahlung in Bar 05.12.2014
    
    'Kredite
    dKredit = 0
    If Not IsNull(rsrs!umsatzkredit) Then
        dKredit = rsrs!umsatzkredit
    Else
        dKredit = 0
    End If

    dSumUms = dSumUms + dKredit
    If dKredit <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Kredit", dKredit, "", 0, 0
    End If
    
    
    
     'VK, nicht Umsatz
    dVKnichtUmsatz = 0
    dVKnichtUmsatz = ermNichtUmsatzinBarBetrag(CByte(gcKasNum))


    dSumBar = dSumBar + dVKnichtUmsatz
    If dVKnichtUmsatz <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "VK, nicht Umsatz", 0, "", dVKnichtUmsatz, 0
    End If
    
    
    
    
    
    
    
    'Gutschein, verkaufte
    dGutschVK = 0
    If Not IsNull(rsrs!vkGutschbar) Then
        dGutschVK = rsrs!vkGutschbar
    Else
        dGutschVK = 0
    End If

    dSumBar = dSumBar + dGutschVK
    If dGutschVK <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Gutschein, verkaufte", 0, "", dGutschVK, 0
    End If
    
    'Gutschein, Umsatz aus
    dGutsch = 0
    If Not IsNull(rsrs!umsatzgutschein) Then
        dGutsch = rsrs!umsatzgutschein
    Else
        dGutsch = 0
    End If

    dSumUms = dSumUms + dGutsch
    If dGutsch <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Gutschein, eingelöste", dGutsch, "", 0, 0
    End If
    
    'Dukaten, Umsatz aus
    dDUKA = 0
    If Not IsNull(rsrs!umsatzDukaten) Then
        dDUKA = rsrs!umsatzDukaten
    Else
        dDUKA = 0
    End If

    dSumUms = dSumUms + dDUKA
    If dDUKA <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Dukaten", dDUKA, "", 0, 0
    End If
    
    'Scheck, Umsatz aus
    dScheck = 0
    If Not IsNull(rsrs!umsatzscheck) Then
        dScheck = rsrs!umsatzscheck
    Else
        dScheck = 0
    End If

    dSumUms = dSumUms + dScheck
    If dScheck <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Scheck", dScheck, "", 0, 0
    End If
    
    'Lastschrift, Umsatz aus
    dLast = 0
    If Not IsNull(rsrs!umsatzlast) Then
        dLast = rsrs!umsatzlast
    Else
        dLast = 0
    End If

    dSumUms = dSumUms + dLast
    If dLast <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Lastschrift", dLast, "", 0, 0
    End If
    
    'Ein und Auszahlungen
    iPos = ermAnzEinAuszahlungAmTag(lDate, CByte(gcKasNum), CByte(gcFilNr), iPos)
    
    'Karte
    dKarte = 0
    If Not IsNull(rsrs!umsatzkarte) Then
        dKarte = rsrs!umsatzkarte
    Else
        dKarte = 0
    End If

    dSumUms = dSumUms + dKarte
    If dKarte <> 0 Then
        iPos = iPos + 1
        InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Karte", dKarte, "", 0, 0
    End If
    
    'Abschöpfung
    dABSCHOPF = 0
    If Not IsNull(rsrs!sumAbschopf) Then
        dABSCHOPF = rsrs!sumAbschopf
    Else
        dABSCHOPF = 0
    End If
    
    dABSCHOPF = -1 * dABSCHOPF

    dSumBar = dSumBar + dABSCHOPF

    iPos = iPos + 1
    InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "Bank", 0, "", dABSCHOPF, 0
    
    iPos = iPos + 1
    InsertKassBuchSatz lDate, iPos, CByte(gcKasNum), CByte(gcFilNr), "gesamt", dSumUms, "", dSumBar, 0
    
End If
rsrs.Close: Set rsrs = Nothing

Screen.MousePointer = 0


Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "InsertAnsKassBuch"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub InsertKassBuchSatz(Insertdate As Long, iPos As Integer, byKasnum As Byte, byFil As Byte, cBezums As String _
, dEURUMS As Double, cBezbar As String, dEURBAR As Double, dEURBANK As Double)
On Error GoTo LOKAL_ERROR

Dim cSQL As String
    
cSQL = "Insert into KABUCH ( "
cSQL = cSQL & " Datum  "
cSQL = cSQL & ", POS  "
cSQL = cSQL & ", BEZUMS  "
cSQL = cSQL & ", EURUMS  "
cSQL = cSQL & ", BEZBAR  "
cSQL = cSQL & ", EURBAR  "
cSQL = cSQL & ", EURBANK  "
cSQL = cSQL & ", KASNUM "
cSQL = cSQL & ", FILIALE  "
cSQL = cSQL & ", SENDOK "
cSQL = cSQL & " ) "
cSQL = cSQL & " values  "
cSQL = cSQL & " ( " & Insertdate & " "
cSQL = cSQL & ", " & iPos & " "
cSQL = cSQL & ", '" & cBezums & "' "
cSQL = cSQL & ", '" & dEURUMS & "' "
cSQL = cSQL & ", '" & cBezbar & "' "
cSQL = cSQL & ", '" & dEURBAR & "' "
cSQL = cSQL & ", '" & dEURBANK & "' "
cSQL = cSQL & ", " & byKasnum & " "
cSQL = cSQL & ", " & byFil & " "
cSQL = cSQL & ", False "
cSQL = cSQL & " ) "
gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "InsertKassBuchSatz"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub InsertWechsel(gdWechsel As Double, byKasnum As Byte)
On Error GoTo LOKAL_ERROR

Dim cSQL As String
    
cSQL = "Insert into Wechsel ( "
cSQL = cSQL & " Datum  "
cSQL = cSQL & ", EURWG  "
cSQL = cSQL & ", KASNUM "
cSQL = cSQL & ", FILIALE  "
cSQL = cSQL & ", SENDOK "
cSQL = cSQL & " ) "
cSQL = cSQL & " values  "
cSQL = cSQL & " ( " & CLng(DateValue(Now)) & " "
cSQL = cSQL & ", '" & gdWechsel & "' "
cSQL = cSQL & ", " & byKasnum & " "
cSQL = cSQL & ", " & gcFilNr & " "
cSQL = cSQL & ", False "
cSQL = cSQL & " ) "
gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "InsertWechsel"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermLastWechselbetrag(byKasnum As Byte) As Double
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim rsrs As Recordset

ermLastWechselbetrag = 0

cSQL = "select autopos, eurwg from Wechsel  "
cSQL = cSQL & " where kasnum = " & byKasnum & " order by autopos desc "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!eurwg) Then
        ermLastWechselbetrag = rsrs!eurwg
    End If
End If
rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermLastWechselbetrag"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermNichtUmsatzinBarBetrag(byKasnum As Byte) As Double
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim rsrs As Recordset

ermNichtUmsatzinBarBetrag = 0

cSQL = "select sum(Betrag)as WERT from NICHTUMSBAR  "
cSQL = cSQL & " where kasnum = " & byKasnum & " and sendok = False "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!Wert) Then
        ermNichtUmsatzinBarBetrag = rsrs!Wert
    End If
End If
rsrs.Close: Set rsrs = Nothing

cSQL = "Update NICHTUMSBAR set sendok = True"
gdBase.Execute cSQL, dbFailOnError

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermNichtUmsatzinBarBetrag"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLastVK() As String
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset
Dim lDatum  As Long

ermLastVK = "keine Verkäufe gefunden"
lDatum = 0

cSQL = "select max(adate)as maxi from Kassjour where adate <= datevalue(now) "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        lDatum = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

If lDatum > 0 Then
    cSQL = "select menge,bezeich,preis,adate,azeit from Kassjour where adate = " & lDatum & " order by azeit desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!ADATE) Then
            ermLastVK = "Datum: " & rsrs!ADATE
        End If
        
        If Not IsNull(rsrs!AZEIT) Then
            ermLastVK = ermLastVK & " Uhrzeit: " & rsrs!AZEIT
        End If
        
        If Not IsNull(rsrs!Menge) Then
            ermLastVK = ermLastVK & " " & rsrs!Menge & "x"
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            ermLastVK = ermLastVK & " " & rsrs!BEZEICH
        End If
        
        If Not IsNull(rsrs!Preis) Then
            ermLastVK = ermLastVK & " Preis: " & Format(rsrs!Preis, "######0.00") & " " & gcWaehrung
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

End If


    
Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermLastVK"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub ermBestMitarbeiter()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim j As Integer
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lAnzNeuKunden As Long
    Dim dAnfangsumsatz As Double
    
    loeschNEW "MITKU" & srechnertab, gdBase
    CreateTable "MITKU" & srechnertab, gdBase
    
    sSQL = "Insert Into MITKU" & srechnertab & " select bednu,bedname from bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("MITKU" & srechnertab)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                lAnzNeuKunden = ermneukundenproMit(rsrs!BEDNU, 14)
                
'                dAnfangsumsatz = ermneukundenUmsproMit(rsrs!BEDNU, 14)
                rsrs.Edit
                rsrs!ANZAHL = lAnzNeuKunden
'                rsrs!anfangums = dAnfangsumsatz
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select top 2 anzahl ,BEDNU ,anfangums from MITKU" & srechnertab & " order by anzahl desc "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
'                lAnzNeuKunden = ermneukundenproMit(rsrs!BEDNU, 14)
                
                dAnfangsumsatz = ermneukundenUmsproMit(rsrs!BEDNU, 14)
                rsrs.Edit
'                rsrs!ANZAHL = lAnzNeuKunden
                rsrs!anfangums = dAnfangsumsatz
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermBestMitarbeiter"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ermBestMitarbeiter14(lblx As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim j As Integer
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lAnzNeuKunden As Long
    Dim dAnfangsumsatz As Double
    
    loeschNEW "MITKU" & srechnertab, gdBase
    CreateTable "MITKU" & srechnertab, gdBase
    
    sSQL = "Insert Into MITKU" & srechnertab & " select bednu,bedname from bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("MITKU" & srechnertab)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                lAnzNeuKunden = ermneukundenproMit(rsrs!BEDNU, 14)
                
                lblx.Caption = rsrs!bedname
                lblx.Refresh
'                dAnfangsumsatz = ermneukundenUmsproMit(rsrs!BEDNU, 14)
                rsrs.Edit
                rsrs!ANZAHL = lAnzNeuKunden
                
                
'                rsrs!anfangums = dAnfangsumsatz
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select  anzahl ,BEDNU ,Bedname,anfangums from MITKU" & srechnertab & " where anzahl > 0 order by anzahl desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
'                lAnzNeuKunden = ermneukundenproMit(rsrs!BEDNU, 14)
                lblx.Caption = rsrs!bedname & " " & rsrs!ANZAHL & " Neukunden"
                lblx.Refresh
                
                dAnfangsumsatz = ermneukundenUmsproMit(rsrs!BEDNU, 14)
                rsrs.Edit
'                rsrs!ANZAHL = lAnzNeuKunden
                rsrs!anfangums = dAnfangsumsatz
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermBestMitarbeiter14"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function inIdentUser(ibednu As Integer) As Boolean
On Error GoTo LOKAL_ERROR

Dim rsrs As Recordset
Dim sSQL As String

inIdentUser = False

sSQL = "Select * from identUser where bednr = " & ibednu
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    inIdentUser = True
End If

rsrs.Close



Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "inIdentUser"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function zeigeBestenMitarbeiterNeukunde() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset

    sSQL = "select top 1 anzahl,bednu,bedname,anfangums from MITKU" & srechnertab & " order by anzahl desc,anfangums desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ANZAHL) Then
            If Val(rsrs!ANZAHL) > 0 Then
                zeigeBestenMitarbeiterNeukunde = "1. Platz "
                zeigeBestenMitarbeiterNeukunde = zeigeBestenMitarbeiterNeukunde & rsrs!bedname
                zeigeBestenMitarbeiterNeukunde = zeigeBestenMitarbeiterNeukunde & Space(45 - Len(zeigeBestenMitarbeiterNeukunde))
                zeigeBestenMitarbeiterNeukunde = zeigeBestenMitarbeiterNeukunde & rsrs!ANZAHL
                zeigeBestenMitarbeiterNeukunde = zeigeBestenMitarbeiterNeukunde & Space(48 - Len(zeigeBestenMitarbeiterNeukunde)) & "Neukunden"
                zeigeBestenMitarbeiterNeukunde = zeigeBestenMitarbeiterNeukunde & Space(60 - Len(zeigeBestenMitarbeiterNeukunde))
                zeigeBestenMitarbeiterNeukunde = zeigeBestenMitarbeiterNeukunde & Format(rsrs!anfangums, "####0.00") & " " & gcWaehrung
            Else
                zeigeBestenMitarbeiterNeukunde = "1. Platz nicht vergeben"
            End If
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "zeigeBestenMitarbeiterNeukunde"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermneukundenproMit(sBednu As String, iTage As Integer) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lAnzNeuKunden As Long
    
    ermneukundenproMit = 0
    
    sSQL = "select count(kundnr) as maxi from kunden where rechnr = " & sBednu
    
    If iTage = 1 Then
    'vor Monat
    
    ElseIf iTage = 2 Then
    'akt Monat
        sSQL = sSQL & " and month(angelegt) = " & Month(Now) & " and year(angelegt) = " & Year(Now)
    Else
        sSQL = sSQL & " and angelegt >= " & CLng(DateValue(Now) - 14)
    End If
   
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermneukundenproMit = Val(rsrs!maxi)
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermneukundenproMit"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermneukundenUmsproMit(sBednu As String, iTage As Integer) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSQL As String
    Dim sSQL1 As String
    Dim rsrs As Recordset
    Dim rsrs1 As Recordset
    Dim lAnzNeuKunden As Long
    
    ermneukundenUmsproMit = 0
    
    sSQL = "select kundnr from kunden where rechnr = " & sBednu
    If iTage = 1 Then
    'vor Monat
    
    ElseIf iTage = 2 Then
    'akt Monat
        sSQL = sSQL & " and angelegt = " & Month(Now) & " and angelegt = " & Year(Now)
    Else
        sSQL = sSQL & " and angelegt >= " & CLng(DateValue(Now) - 14)
    End If
   
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            sSQL1 = "select sum(Preis)as maxi from kassjour where kundnr = " & rsrs!Kundnr
            If iTage = 1 Then
            'vor Monat
            
            ElseIf iTage = 2 Then
            'akt Monat
                sSQL1 = sSQL1 & " and adate = " & Month(Now) & " and adate = " & Year(Now)
            Else
                sSQL1 = sSQL1 & " and adate >= " & CLng(DateValue(Now) - 14)
            End If
            sSQL1 = sSQL1 & " and ums_ok = 'J'"
            Set rsrs1 = gdBase.OpenRecordset(sSQL1)
            If Not rsrs1.EOF Then
                If Not IsNull(rsrs1!maxi) Then
                    ermneukundenUmsproMit = ermneukundenUmsproMit + CDbl((rsrs1!maxi))
                End If
            End If
            rsrs1.Close: Set rsrs1 = Nothing
            
            'Teil 2
            
            sSQL1 = "select sum(Preis)as maxi from kundkass where kundnr = " & rsrs!Kundnr
            If iTage = 1 Then
            'vor Monat
            
            ElseIf iTage = 2 Then
            'akt Monat
                sSQL1 = sSQL1 & " and adate = " & Month(Now) & " and angelegt = " & Year(Now)
            Else
                sSQL1 = sSQL1 & " and adate >= " & CLng(DateValue(Now) - 14)
            End If
'            sSQL1 = sSQL1 & " and ums_ok = 'J'"
            Set rsrs1 = gdBase.OpenRecordset(sSQL1)
            If Not rsrs1.EOF Then
                If Not IsNull(rsrs1!maxi) Then
                    ermneukundenUmsproMit = ermneukundenUmsproMit + CDbl((rsrs1!maxi))
                End If
            End If
            rsrs1.Close: Set rsrs1 = Nothing
            
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermneukundenproMit"
    Fehler.gsFehlertext = "Im Programmteil Neukundenauswertung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function rechneNeuKunden()
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset
Dim lAnz    As Long


loeschNEW "NEUKU" & srechnertab, gdBase
CreateTable "NEUKU" & srechnertab, gdBase

lAnz = 0

cSQL = "select count(kundnr)as maxi from Kunden "
cSQL = cSQL & " where year(angelegt) = year(datevalue(now)) "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        lAnz = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

cSQL = "Insert into NEUKU" & srechnertab & " (Anzahl,Thema) values (" & lAnz & ",'akt Jahr')"
gdBase.Execute cSQL, dbFailOnError

lAnz = 0

cSQL = "select count(kundnr)as maxi from Kunden "
cSQL = cSQL & " where year(angelegt) = year(datevalue(now))-1 "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        lAnz = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

cSQL = "Insert into NEUKU" & srechnertab & " (Anzahl,Thema) values (" & lAnz & ",'vj Jahr')"
gdBase.Execute cSQL, dbFailOnError

lAnz = 0

cSQL = "select count(kundnr)as maxi from Kunden "
cSQL = cSQL & " where month(angelegt) = month(datevalue(now)) "
cSQL = cSQL & " and year(angelegt) = year(datevalue(now)) "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        lAnz = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

cSQL = "Insert into NEUKU" & srechnertab & " (Anzahl,Thema) values (" & lAnz & ",'akt Monat')"
gdBase.Execute cSQL, dbFailOnError

lAnz = 0

cSQL = "select count(kundnr)as maxi from Kunden "
cSQL = cSQL & " where month(angelegt) = month(datevalue(now)) "
cSQL = cSQL & " and year(angelegt) = year(datevalue(now))-1 "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        lAnz = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

cSQL = "Insert into NEUKU" & srechnertab & " (Anzahl,Thema) values (" & lAnz & ",'vj akt Monat')"
gdBase.Execute cSQL, dbFailOnError

lAnz = 0

cSQL = "select count(kundnr)as maxi from Kunden "
cSQL = cSQL & " where angelegt = datevalue(now)-1 "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        lAnz = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

cSQL = "Insert into NEUKU" & srechnertab & " (Anzahl,Thema) values (" & lAnz & ",'gestern')"
gdBase.Execute cSQL, dbFailOnError


Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "rechneNeuKunden"
    Fehler.gsFehlertext = "Im Programmteil Neukunden ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermNeuKunden(iArt As Integer) As String
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset
Dim lAnz    As Long
Dim cThema  As String

ermNeuKunden = ""
lAnz = 0

Select Case iArt
    Case 1
        cThema = "akt Jahr"
    Case 2
        cThema = "akt Monat"
    Case 3
        cThema = "vj Jahr"
    Case 4
        cThema = "vj akt Monat"
    Case 5
        cThema = "gestern"
End Select

cSQL = "select anzahl from NEUKU" & srechnertab & " where thema =  '" & cThema & "'"

Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!ANZAHL) Then
        lAnz = rsrs!ANZAHL
    End If
End If
rsrs.Close: Set rsrs = Nothing

If iArt = 1 Then
    ermNeuKunden = "gesamt " & Year(Now) & ":"
    ermNeuKunden = ermNeuKunden & Space(13 - Len(ermNeuKunden)) & lAnz
ElseIf iArt = 2 Then
    ermNeuKunden = Format(DateValue(Now), "mmm") & " " & Year(Now) & ":"
    ermNeuKunden = ermNeuKunden & Space(10 - Len(ermNeuKunden)) & lAnz
ElseIf iArt = 3 Then
    ermNeuKunden = "gesamt " & Year(Now) - 1 & ":"
    ermNeuKunden = ermNeuKunden & Space(13 - Len(ermNeuKunden)) & lAnz
ElseIf iArt = 4 Then
    ermNeuKunden = Format(DateValue(Now), "mmm") & " " & Year(Now) - 1 & ":"
    ermNeuKunden = ermNeuKunden & Space(10 - Len(ermNeuKunden)) & lAnz
ElseIf iArt = 5 Then
    ermNeuKunden = "Gestern " & Format(DateValue(Now) - 1, "dd.mm.") & ":"
    ermNeuKunden = ermNeuKunden & Space(16 - Len(ermNeuKunden)) & lAnz
End If

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermNeuKunden"
    Fehler.gsFehlertext = "Im Programmteil Neukunden ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermAnzEinAuszahlungAmTag(Insertdate As Long, byKasnum As Byte, byFil As Byte, iPos As Integer) As Integer
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset
Dim dBetrag As Double

ermAnzEinAuszahlungAmTag = iPos

sSQL = "Select Bezeich,Betrag   from EINAUSKB where adate = " & Insertdate
sSQL = sSQL & " and kasnum = " & byKasnum
sSQL = sSQL & " and sendok =  False "
'sSQL = sSQL & " and Filiale = " & byFil
sSQL = sSQL & " and Art = 'EINZAHLUNG'"

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    Do While Not rsrs.EOF
    
        dBetrag = 0
        If Not IsNull(rsrs!Betrag) Then
            dBetrag = rsrs!Betrag
        Else
            dBetrag = 0
        End If
        
        dSumBar = dSumBar + dBetrag

        ermAnzEinAuszahlungAmTag = ermAnzEinAuszahlungAmTag + 1
        InsertKassBuchSatz Insertdate, ermAnzEinAuszahlungAmTag, CByte(gcKasNum), CByte(gcFilNr), "Einzahlung " & rsrs!BEZEICH, 0, "", dBetrag, 0
        
    rsrs.MoveNext
    Loop
    
End If
rsrs.Close: Set rsrs = Nothing

sSQL = "Select Bezeich,(-1 * Betrag) as Wert from EINAUSKB where adate = " & Insertdate
sSQL = sSQL & " and kasnum = " & byKasnum
sSQL = sSQL & " and sendok =  False "
'sSQL = sSQL & " and Filiale = " & byFil
sSQL = sSQL & " and Art = 'AUSZAHLUNG'"

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    Do While Not rsrs.EOF
    
        dBetrag = 0
        If Not IsNull(rsrs!Wert) Then
            dBetrag = rsrs!Wert
        Else
            dBetrag = 0
        End If
        
        dSumBar = dSumBar + dBetrag

        ermAnzEinAuszahlungAmTag = ermAnzEinAuszahlungAmTag + 1
        InsertKassBuchSatz Insertdate, ermAnzEinAuszahlungAmTag, CByte(gcKasNum), CByte(gcFilNr), "Auszahlung " & rsrs!BEZEICH, 0, "", dBetrag, 0
        
    rsrs.MoveNext
    Loop
    
End If
rsrs.Close: Set rsrs = Nothing



Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermAnzEinAuszahlungAmTag"
    Fehler.gsFehlertext = "Im Programmteil Kassenbuch ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermWertrabattf_Artikel() As Double
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim dWert       As Double
    Dim dKdBonus    As Double
    Dim dKJPreis    As Double
    Dim cArtNr      As String
    
    ermWertrabattf_Artikel = 0
    
    lAnzSatz = frmWKL20.List1.ListCount
    dKdBonus = 0
    
    For lAktSatz = 0 To lAnzSatz - 1
        dKJPreis = 0
        cLBSatz = frmWKL20.List1.list(lAktSatz)
        
        cArtNr = Mid(cLBSatz, 7, 6)
        
        If cArtNr <> "666666" Then
        
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            dKJPreis = Val(ctmp)
            
            If Mid(cLBSatz, 6, 1) <> "*" Then
                dKdBonus = dKdBonus + dKJPreis
            End If
        End If
    Next lAktSatz

    ermWertrabattf_Artikel = dKdBonus
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermWertrabattf_Artikel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermaktKassensoll() As Double
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsRs3 As Recordset
    Dim dWert As Double
    Dim dSumme As Double
    Dim dUmsatz As Double
    Dim cMWSK As String
    Dim dEinzahlung As Double
    Dim dAuszahlung As Double
    Dim dAuszGutsch As Double
    Dim dBar As Double
    Dim dKunden As Double
    Dim dScheck As Double
    Dim dKredit As Double
    Dim dKarte As Double
    Dim dLast As Double
    Dim dUmsBar As Double
    Dim dUmsScheck As Double

    Dim dZhlgGutsch As Double

    Dim dKasse As Double
    Dim dKassenBargeld As Double
    Dim dKassenSchecks As Double
    Dim dSchVerkauf As Double
    Dim dBarVerkauf As Double

    Dim dGutschein As Double
    Dim dGutschBar As Double
    Dim dGutschSch As Double
    Dim dGutschKre As Double
    Dim dGutschKar As Double
    Dim dGutschLast As Double
    Dim dGutschGUTSCH As Double
    Dim dABSCHOPF As Double
    Dim dKDIFF As Double
    Dim dTDIFF As Double
    Dim dDUKA As Double
    Dim dWECHSEL As Double
    Dim dEinrGutsch As Double

    Dim dTilgung As Double
    Dim dTilgBar As Double
    Dim dTilgSch As Double
    Dim dTilgGut As Double
    Dim dTilgKar As Double
    
    ermaktKassensoll = 0
    
    cSQL = "Select KASNUM"
    cSQL = cSQL & ", SUM(UMS_BAR) as SUMS_BAR"
    cSQL = cSQL & ", SUM(UMS_KRED) as SUMS_KRED"
    cSQL = cSQL & ", SUM(UMS_KARTE) as SUMS_KARTE"
    cSQL = cSQL & ", SUM(UMS_SCHECK) as SUMS_SCHEC"
    cSQL = cSQL & ", SUM(UMS_LAST) as SUMS_LAST"
    
    cSQL = cSQL & ", SUM(TILGBAR) as STILGBAR"
    cSQL = cSQL & ", SUM(TILGSCH) as STILGSCH"
    cSQL = cSQL & ", SUM(TILGGUT) as STILGGUT"
    cSQL = cSQL & ", SUM(TILGKAR) as STILGKAR"
    
    cSQL = cSQL & ", SUM(GUTSCHBAR) as SGUTSCHBAR"
    cSQL = cSQL & ", SUM(GUTSCHSCH) as SGUTSCHSCH"
    cSQL = cSQL & ", SUM(GUTSCHKRE) as SGUTSCHKRE"
    cSQL = cSQL & ", SUM(GUTSCHKAR) as SGUTSCHKAR"
    cSQL = cSQL & ", SUM(GUTSCHLAST) as SGUTSCHLAS"
    cSQL = cSQL & ", SUM(GUTSCHGUTSCH) as SGUTSCHGUTSCH"
    cSQL = cSQL & ", SUM(ABSCHOPF) as SABSCHOPF"
    cSQL = cSQL & ", SUM(KDIFF) as SKDIFF"
    cSQL = cSQL & ", SUM(TDIFF) as STDIFF"
    cSQL = cSQL & ", SUM(DUKA) as SDUKA"
    cSQL = cSQL & ", SUM(WECHSEL) as SWECHSEL"
    
    cSQL = cSQL & ", SUM(BARVERKAUF) as SBARVERKAU"
    cSQL = cSQL & ", SUM(SCHVERKAUF) as SSCHVERKAU"
    
    cSQL = cSQL & ", SUM(AUSZAHLUNG) as SAUSZAHLUN"
    cSQL = cSQL & ", SUM(EINZAHLUNG) as SEINZAHLUN"
    cSQL = cSQL & ", SUM(AUSZGUTSCH) as SAUSZGUTSC"
    
    cSQL = cSQL & ", SUM(SPREIS_GES) as SSPREIS_GE"
    cSQL = cSQL & ", SUM(SPREIS_ANZ) as SSPREIS_AN"
    cSQL = cSQL & ", SUM(GESRAB_GES) as SGESRAB_GE"
    cSQL = cSQL & ", SUM(GESRAB_ANZ) as SGESRAB_AN"
    cSQL = cSQL & ", SUM(ARTRAB_GES) as SARTRAB_GE"
    cSQL = cSQL & ", SUM(ARTRAB_ANZ) as SARTRAB_AN"
    cSQL = cSQL & ", SUM(STORNO_GES) as SSTORNO_GE"
    cSQL = cSQL & ", SUM(STORNO_ANZ) as SSTORNO_AN"
    
    cSQL = cSQL & ", SUM(ZHLGGUTSCH) as SZHLGGUTSC"
    cSQL = cSQL & ", SUM(KUNDENZAHL) as SKUNDENZAH"
    cSQL = cSQL & ", SUM(GELDFACH) as SGELDFACH"
    
    cSQL = cSQL & ", SUM(EINRGUTSCH) as SEINRGUTSC"
    cSQL = cSQL & ", SUM(RESTGUTSCH) as SRESTGUTSC"
    cSQL = cSQL & ", SUM(GUTSCHEIN) as SGUTSCH"
    cSQL = cSQL & " from AFCSTAT where KASNUM = " & gcKasNum & " group by KASNUM "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        dWert = rsrs.RecordCount
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SUMS_BAR) Then
            dWert = rsrs!SUMS_BAR
        Else
            dWert = 0
        End If
        dUmsBar = dWert
        
        If Not IsNull(rsrs!SUMS_SCHEC) Then
            dWert = rsrs!SUMS_SCHEC
        Else
            dWert = 0
        End If
        dUmsScheck = dWert
        
        If Not IsNull(rsrs!SUMS_KARTE) Then
            dWert = rsrs!SUMS_KARTE
        Else
            dWert = 0
        End If
        dKarte = dWert
                
        If Not IsNull(rsrs!SUMS_KRED) Then
            dWert = rsrs!SUMS_KRED
        Else
            dWert = 0
        End If
        dKredit = dWert
        
        
        If Not IsNull(rsrs!SUMS_LAST) Then
            dWert = rsrs!SUMS_LAST
        Else
            dWert = 0
        End If
        dLast = dWert
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        
        
        If Not IsNull(rsrs!STILGBAR) Then
            dWert = rsrs!STILGBAR
        Else
            dWert = 0
        End If
        dTilgBar = dWert
        
        If Not IsNull(rsrs!STILGSCH) Then
            dWert = rsrs!STILGSCH
        Else
            dWert = 0
        End If
        dTilgSch = dWert
        
        If Not IsNull(rsrs!STILGGUT) Then
            dWert = rsrs!STILGGUT
        Else
            dWert = 0
        End If
        dTilgGut = dWert
        
        If Not IsNull(rsrs!STILGKAR) Then
            dWert = rsrs!STILGKAR
        Else
            dWert = 0
        End If
        dTilgKar = dWert
        
        dTilgung = dTilgBar + dTilgSch + dTilgGut + dTilgKar
        
        If Not IsNull(rsrs!SGUTSCHBAR) Then
            dWert = rsrs!SGUTSCHBAR
        Else
            dWert = 0
        End If
        dGutschBar = dWert
        
        If Not IsNull(rsrs!SGUTSCHSCH) Then
            dWert = rsrs!SGUTSCHSCH
        Else
            dWert = 0
        End If
        dGutschSch = dWert
        
        If Not IsNull(rsrs!SGUTSCHKRE) Then
            dWert = rsrs!SGUTSCHKRE
        Else
            dWert = 0
        End If
        dGutschKre = dWert
        
        If Not IsNull(rsrs!SGUTSCHKAR) Then
            dWert = rsrs!SGUTSCHKAR
        Else
            dWert = 0
        End If
        dGutschKar = dWert
        
        If Not IsNull(rsrs!SGUTSCHLAS) Then
            dWert = rsrs!SGUTSCHLAS
        Else
            dWert = 0
        End If
        dGutschLast = dWert
        
        If Not IsNull(rsrs!SGUTSCHGUTSCH) Then
            dWert = rsrs!SGUTSCHGUTSCH
        Else
            dWert = 0
        End If
        dGutschGUTSCH = dWert
        
        If Not IsNull(rsrs!SABSCHOPF) Then
            dWert = rsrs!SABSCHOPF
        Else
            dWert = 0
        End If
        dABSCHOPF = dWert
        

        dTDIFF = 0
        
        If Not IsNull(rsrs!SDUKA) Then
            dWert = rsrs!SDUKA
        Else
            dWert = 0
        End If
        dDUKA = dWert
        
        If Not IsNull(rsrs!SWECHSEL) Then
            dWert = rsrs!SWECHSEL
        Else
            dWert = 0
        End If
        dWECHSEL = dWert
        
        If Not IsNull(rsrs!sGutsch) Then
            dWert = rsrs!sGutsch
        Else
            dWert = 0
        End If
        dGutschein = dWert
        
        If Not IsNull(rsrs!SSCHVERKAU) Then
            dWert = rsrs!SSCHVERKAU
        Else
            dWert = 0
        End If
        dSchVerkauf = dWert
        
        dKassenSchecks = dSchVerkauf + dGutschSch + dTilgSch
        
        If Not IsNull(rsrs!SAUSZAHLUN) Then
            dWert = rsrs!SAUSZAHLUN
        Else
            dWert = 0
        End If
        dAuszahlung = dWert
        
        If Not IsNull(rsrs!SEINZAHLUN) Then
            dWert = rsrs!SEINZAHLUN
        Else
            dWert = 0
        End If
        dEinzahlung = dWert
        
        If Not IsNull(rsrs!SAUSZGUTSC) Then
            dWert = rsrs!SAUSZGUTSC
        Else
            dWert = 0
        End If
        dAuszGutsch = dWert
        
        If Not IsNull(rsrs!SBARVERKAU) Then
            dWert = rsrs!SBARVERKAU
        Else
            dWert = 0
        End If
        dBarVerkauf = dWert
        
        dKassenBargeld = dBarVerkauf + dGutschBar + dTilgBar + dEinzahlung - dAuszahlung - dAuszGutsch - dABSCHOPF + dWECHSEL
    
        ermaktKassensoll = dKassenBargeld '+ dKassenSchecks ohne Schecks Hickmann
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermaktKassensoll"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermaktWechselgeld() As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    Dim db As DAO.Database
    
    Dim cPfad       As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = ShortPath(cPfad)
    
    Set db = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    ermaktWechselgeld = 0
    
    cSQL = "Select KASNUM"
    cSQL = cSQL & ", SUM(WECHSEL) as SWECHSEL"
    cSQL = cSQL & " from AFCSTAT where KASNUM = " & gcKasNum & " group by KASNUM "
    Set rsrs = db.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!SWECHSEL) Then
            ermaktWechselgeld = rsrs!SWECHSEL
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    db.Close
    
Exit Function
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "ermaktWechselgeld"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function GibtesdiesenArtikel(cART As String) As Boolean
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

GibtesdiesenArtikel = False

sSQL = "Select * from Artikel where artnr = " & cART
Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    GibtesdiesenArtikel = True
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul20"
    Fehler.gsFunktion = "GibtesdiesenArtikel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub schreibe_CouponCSV(sBudnr As String, lAuswerttag As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL                As String
    Dim cPfad               As String
    Dim cdatei              As String
    Dim cPfad1              As String
    Dim iRet                As Integer
    Dim rsrs                As Recordset
    Dim sAusgabedatname     As String
    Dim iFileNr             As Integer
    Dim lPos                As Long
    Dim cSatz               As String
    Dim cFeld               As String
    Dim dGeld               As Double
    
    Screen.MousePointer = 11
    
'    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    cPfad1 = App.Path
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    cPfad1 = cPfad1 & "EDI\"
''''    Kill cPfad1 & "*.*"
    
    sSQL = "Select * from COUPONPRINT"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        sAusgabedatname = "dronova_" & sBudnr & "_" & Format(lAuswerttag, "yyyymmdd") & "_" & Format(TimeValue(Now), "HHMM") & ".csv"

        cPfad1 = App.Path
        If Right(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "EDI\" & sAusgabedatname
        cPfad = cPfad1 & "EDI"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
   
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = "DRONOVA;DRONOVA;" & sBudnr & ";"
            
            cFeld = ""
            If Not IsNull(rsrs!ADATE) Then
                cFeld = Format(rsrs!ADATE, "DD.MM.YYYY")
            End If
            cSatz = cSatz & cFeld & ";"
            
            cFeld = ""
            If Not IsNull(rsrs!EAN) Then
                cFeld = rsrs!EAN
            End If
            cSatz = cSatz & cFeld & ";"
            
            cFeld = ""
            If Not IsNull(rsrs!Menge) Then
                cFeld = rsrs!Menge
            End If
            cSatz = cSatz & cFeld & ";"
            
            If Not IsNull(rsrs!Preis) Then
                cFeld = Format(rsrs!Preis, "######0.00")
            End If
            cFeld = SwapStr(cFeld, ",", ".")
            cSatz = cSatz & cFeld
            
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul20"
        Fehler.gsFunktion = "schreibe_CouponCSV"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
