Attribute VB_Name = "mdlEtikett"
Option Explicit
Public Sub SicherInEtisic(sArtnr As String, sFil As String, checkX As CheckBox)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Insert into Etisic select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", VKPR "
    cSQL = cSQL & ", BESTAND "
    cSQL = cSQL & ", ANZAHL "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", Pcname "
    cSQL = cSQL & ", FILNR "
    
    cSQL = cSQL & ", '" & srechnertab & "' as  DelPcname "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as DelDATE  "
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as DelTIME  "
    
    cSQL = cSQL & " from Etidru "
    cSQL = cSQL & " where Artnr = " & sArtnr
    cSQL = cSQL & " and FILNR = " & sFil
    
    If checkX.Value = vbChecked Then
        cSQL = cSQL & " and PCNAME = '" & srechnertab & "' "
    End If
    
    gdBase.Execute cSQL, dbFailOnError
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "SicherInEtisic"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
'    Resume Next
End Sub
Public Sub SicherInEtisicALL(checkX As CheckBox)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Insert into Etisic select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", VKPR "
    cSQL = cSQL & ", BESTAND "
    cSQL = cSQL & ", ANZAHL "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", Pcname "
    cSQL = cSQL & ", FILNR "
    
    cSQL = cSQL & ", '" & srechnertab & "' as  DelPcname "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as DelDATE  "
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as DelTIME  "
    
    cSQL = cSQL & " from Etidru "
    
    If checkX.Value = vbChecked Then
        cSQL = cSQL & " where PCNAME = '" & srechnertab & "' "
    End If
    
    gdBase.Execute cSQL, dbFailOnError
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "SicherInEtisicALL"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub ZurückSicherInEtidru(sArtnr As String, sFil As String)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Insert into Etidru select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", VKPR "
    cSQL = cSQL & ", BESTAND "
    cSQL = cSQL & ", ANZAHL "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", Pcname "
    cSQL = cSQL & ", FILNR "
    cSQL = cSQL & " from Etisic "
    cSQL = cSQL & " where Artnr = " & sArtnr
    cSQL = cSQL & " and FILNR = " & sFil
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from ETISIC where ARTNR = " & sArtnr
    cSQL = cSQL & " and FILNR = " & sFil
    gdBase.Execute cSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "ZurückSicherInEtidru"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub ZurückSicherInEtidruALL()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    cSQL = "Insert into Etidru select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & ", BEZEICH "
    cSQL = cSQL & ", VKPR "
    cSQL = cSQL & ", BESTAND "
    cSQL = cSQL & ", ANZAHL "
    cSQL = cSQL & ", LIBESNR "
    cSQL = cSQL & ", EAN "
    cSQL = cSQL & ", LINR "
    cSQL = cSQL & ", LPZ "
    cSQL = cSQL & ", Pcname "
    cSQL = cSQL & ", FILNR "
    cSQL = cSQL & " from Etisic "
    gdBase.Execute cSQL, dbFailOnError
    
    gdBase.Execute "Delete from ETISic", dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "ZurückSicherInEtidruALL"
    Fehler.gsFehlertext = "Im Programmteil Bestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub CreateArtikelsic()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "Artikelsic", gdBase
    
    cSQL = "Select * into Artikelsic from Artikel where Artnr = -1"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table Artikelsic add DelPcname varchar(30) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table Artikelsic add DelDATE DATETIME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table Artikelsic add DelTIME varchar(10) "
    gdBase.Execute cSQL, dbFailOnError
    
    

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "CreateArtikelsic"
    Fehler.gsFehlertext = "Im Programmteil Artikel löschen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub CreateArtliefsic()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    
    loeschNEW "Artliefsic", gdBase
    
    cSQL = "Select * into Artliefsic from Artlief where Artnr = -1"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table Artliefsic add DelPcname varchar(30) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table Artliefsic add DelDATE DATETIME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " Alter table Artliefsic add DelTIME varchar(10) "
    gdBase.Execute cSQL, dbFailOnError

    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "CreateArtliefsic"
    Fehler.gsFehlertext = "Im Programmteil Artikel löschen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub SicherInArtikelsic(lartnr As Long)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    If NewTableSuchenDBKombi("Artikelsic", gdBase) = False Then
        CreateArtikelsic
    End If
    
     If NewTableSuchenDBKombi("Artliefsic", gdBase) = False Then
        CreateArtliefsic
    End If
    
    cSQL = "Insert into Artikelsic select Artikel.* "
    cSQL = cSQL & ", '" & srechnertab & "' as  DelPcname "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as DelDATE  "
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as DelTIME  "
    cSQL = cSQL & " from Artikel "
    cSQL = cSQL & " where Artnr = " & lartnr
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Artliefsic select Artlief.* "
    cSQL = cSQL & ", '" & srechnertab & "' as  DelPcname "
    cSQL = cSQL & ", '" & DateValue(Now) & "' as DelDATE  "
    cSQL = cSQL & ", '" & TimeValue(Now) & "' as DelTIME  "
    cSQL = cSQL & " from Artlief "
    cSQL = cSQL & " where Artnr = " & lartnr
    gdBase.Execute cSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "SicherInArtikelsic"
    Fehler.gsFehlertext = "Im Programmteil Artikel löschen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub ZurückSicherInArtikel(sArtnr As String)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    
    loeschNEW "Artikelttempsic", gdBase
    
    cSQL = "Select * into Artikelttempsic from ArtikelSic "
    cSQL = cSQL & " where Artnr = " & sArtnr
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artikelttempsic DROP DelPcname "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artikelttempsic DROP DelDATE "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artikelttempsic DROP DelTIME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Artikel Select * from Artikelttempsic "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "Artliefttempsic", gdBase
    
    cSQL = "Select * into Artliefttempsic from ArtliefSic "
    cSQL = cSQL & " where Artnr = " & sArtnr
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artliefttempsic DROP DelPcname "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artliefttempsic DROP DelDATE "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artliefttempsic DROP DelTIME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Artlief Select *  from Artliefttempsic "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    cSQL = "Delete from ArtikelSIC where ARTNR = " & sArtnr
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from ArtliefSIC where ARTNR = " & sArtnr
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "Artliefttempsic", gdBase
    loeschNEW "Artikelttempsic", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "ZurückSicherInArtikel"
    Fehler.gsFehlertext = "Im Programmteil Artikel wiederherstellen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub ZurückSicherInArtikelALL()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "Artikelttempsic", gdBase
    
    cSQL = "Select * into Artikelttempsic from ArtikelSic "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artikelttempsic DROP DelPcname "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artikelttempsic DROP DelDATE "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artikelttempsic DROP DelTIME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Artikel Select * from Artikelttempsic "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "Artliefttempsic", gdBase
    
    cSQL = "Select * into Artliefttempsic from ArtliefSic "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artliefttempsic DROP DelPcname "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artliefttempsic DROP DelDATE "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Alter table Artliefttempsic DROP DelTIME "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into Artlief Select * from Artliefttempsic "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    gdBase.Execute "Delete from ArtikelSic", dbFailOnError
    gdBase.Execute "Delete from ArtliefSic", dbFailOnError
    loeschNEW "Artliefttempsic", gdBase
    loeschNEW "Artikelttempsic", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEtikett"
    Fehler.gsFunktion = "ZurückSicherInArtikelALL"
    Fehler.gsFehlertext = "Im Programmteil Artikel wiederherstellen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
