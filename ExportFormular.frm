VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ExportFormular 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   114032641
      CurrentDate     =   44400
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   114032641
      CurrentDate     =   44400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ausgabe Pfad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Starten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7080
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox ChkOeffnen 
      Caption         =   "die exportierten Dateien am Prozess-Ende automatisch im Ordner öffnen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9735
   End
   Begin VB.Label lblProgress 
      Caption         =   "*****"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   13935
   End
   Begin VB.Label Label1 
      Caption         =   "Datei :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblDatei 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   8295
   End
End
Attribute VB_Name = "ExportFormular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AutomatischDateiOffnen As Boolean
Dim rsRes As Recordset
Dim iDsF As Integer
Dim DateiZuErstellen As String

Dim WirdVerarbeitet As Boolean
     

Private Sub ChkOeffnen_Click()
 
 
 If ChkOeffnen.value = vbChecked Then
    
          AutomatischDateiOffnen = True
     Else
          AutomatischDateiOffnen = False
    
 End If
 
 
 
End Sub

Private Sub Command1_Click()

ChooseFile.Top = Me.Top - 200
ChooseFile.Left = Me.Left + Me.Width / 4
ChooseFile.Show 1

End Sub

Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR

If Trim(gbDsFinvkPfad) = "" Then

 MsgBox ("Bitte erstmal Pfad wählen ! ! !")
 
 Else
   
     
    Command2.Enabled = False
     
    If WirdVerarbeitet Then
       
       WirdVerarbeitet = False
       alteTabellenVonDsFinvKLoschen
       StarteExportieren
       
    End If
     
End If

Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil DsFinvK Expo. ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub


Sub StarteExportieren()
On Error GoTo LOKAL_ERROR
    
      ' Hinweis: die Tabelle [ ToDsFinvK ] ist eine temp Tabelle, die für die zwischen-Prozessierung benutzt wird
      ' wenn die Tabelle [ ToDsFinvK ] existiert, loesch die, um von Grund auf neu zu starten
      If NewTableSuchenDB("ToDsFinvK", gdBase) Then
      
        sSQL = "drop Table ToDsFinvK"
        gdBase.Execute sSQL, dbFailOnError
        
      End If
      
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''' START Einzelaufzeichnungsmodul '''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
                   
            '**************************************************************************************> Bonkopf.csv
                                                    lblDatei.Caption = "Bonkopf.csv"
                                                    lblDatei.Refresh
               If NewTableSuchenDB("Bonkopf_tmp", gdBase) Then
                   sSQL = "drop Table Bonkopf_tmp"
                   gdBase.Execute sSQL, dbFailOnError
               End If
  
             '1. grunde Spalten abfragen (Hinweis: (BELEGNR) ist einfach BON_NR wie z.b 1001,1002. aber (BON_ID) ist eine fortlaufende Zahl, die jetzt erstellt wird
             lblProgress.Caption = "1. die entsprechenden Spalten für Bonkopf werden ermittelt ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID, DATUM , UHRZEIT ,'' as Z_NR,BONNR as BON_NR,'Beleg' as BON_TYP,'' as BON_NAME,'' as TERMINAL_ID,'' as BON_STORNO,'' as BON_START,'' as BON_ENDE,'' as BEDIENER_ID,'' as BEDIENER_NAME,BETRAG as UMS_BRUTTO,'' as KUNDE_NAME,KUNDNR as KUNDE_ID,'' as KUNDE_TYP,'' as KUNDE_STRASSE,'' as KUNDE_PLZ,'' as KUNDE_ORT,'' as KUNDE_LAND,'' as KUNDE_USTID,'' as BON_NOTIZ INTO Bonkopf_tmp FROM KASSBON WHERE Datevalue(DATUM) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "')")
             '2. fortlaufende BON_ID genarieren"
             lblProgress.Caption = "2. fortlaufende BON_ID genarieren ..."
             lblProgress.Refresh
             gdBase.Execute ("ALTER TABLE Bonkopf_tmp ADD COLUMN BON_ID COUNTER (0,1)")
             '3. Bediener-Daten aller Belege ermitteln ..."
             lblProgress.Caption = "3. Bediener-info aller Belege ermitteln ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE Bonkopf_tmp BK INNER JOIN KASSJOUR KJ ON BK.DATUM=KJ.ADATE AND BK.UHRZEIT=KJ.AZEIT AND BK.Z_KASSE_ID=KJ.KASNUM AND BK.BON_NR=KJ.BELEGNR SET BK.BEDIENER_ID=KJ.Bediener")
             gdBase.Execute ("UPDATE Bonkopf_tmp BK INNER JOIN BEDNAME BN ON VAL(BK.BEDIENER_ID)=BN.BEDNU SET BK.BEDIENER_NAME=BN.BEDNAME")
             '4. Kunden-Info ermitteln ..."
             lblProgress.Caption = "4. Kunden-Info ermitteln ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE Bonkopf_tmp BK INNER JOIN KUNDEN KU ON BK.KUNDE_ID=KU.KUNDNR SET BK.KUNDE_NAME=KU.VORNAME & ' ' & KU.NAME,BK.KUNDE_STRASSE=KU.STRASSE,BK.KUNDE_PLZ=KU.PLZ,BK.KUNDE_ORT=KU.STADT")
             gdBase.Execute ("SELECT Z_KASSE_ID,Format (DATUM & ' ' & UHRZEIT, 'yyyy-mm-dd\Thh:nn:ss') as Z_ERSTELLUNG,Z_NR,BON_ID,BON_NR,BON_TYP,BON_NAME,TERMINAL_ID,BON_STORNO,BON_START,BON_ENDE,BEDIENER_ID,BEDIENER_NAME,UMS_BRUTTO,KUNDE_NAME,KUNDE_ID,KUNDE_TYP,KUNDE_STRASSE,KUNDE_PLZ,KUNDE_ORT,KUNDE_LAND,KUNDE_USTID,BON_NOTIZ INTO Bonkopf FROM Bonkopf_tmp ")
             gdBase.Execute ("DROP TABLE Bonkopf_tmp")
             
             
          
            '**************************************************************************************> Bonpos.csv
                                                    lblDatei.Caption = "Bonpos.csv"
                                                    lblDatei.Refresh
            
            '1. die grundlegende Columns aus der Tabelle [Kassjour] in der Tabelle [ToDsFinvK] kopieren ( nur für das gegenwärtige jahr )
            '   Hinweis: VKPR ist BruttoPreis
            lblProgress.Caption = "1. die grundlegende Columns aus der Tabelle [Kassjour] in der Tabelle [ToDsFinvK] kopieren  ..."
            lblProgress.Refresh
            gdBase.Execute ("SELECT KASNUM,ADATE,AZEIT,BELEGNR,'' as BON_ID,BEZEICH,ARTNR,EAN,AGN,MENGE,VKPR,PREIS,MWST,'' as Z_NR,'' as POS_TERMINAL_ID,'Umsatz' as GV_TYP,'' as GV_NAME,'' as INHAUS,'' as AGENTUR_ID,'Stueck' as EINHEIT,'' as FAKTOR,'' as P_STORNO,'' as WARENGR INTO ToDsFinvK FROM Kassjour WHERE Datevalue(ADATE) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "')")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN POS_ZEILE COUNTER (0,1)")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN GUTSCHEIN_NR NUMBER")
            'ToDsFinvK & Bonkopf sind miteinander durch [ BON_ID ] verknüpfen
            'gdBase.Execute ("UPDATE ToDsFinvK TD INNER JOIN Bonkopf BK ON BK.Z_KASSE_ID=TD.KASNUM AND BK.Z_ERSTELLUNG=Format (TD.ADATE & ' ' & TD.AZEIT, 'yyyy-mm-dd\Thh:nn:ss') AND BK.BON_NR=TD.BELEGNR SET TD.BON_ID=BK.BON_ID")
            gdBase.Execute ("UPDATE ToDsFinvK TD INNER JOIN Bonkopf BK ON BK.Z_KASSE_ID=TD.KASNUM AND DateValue(LEFT(BK.Z_ERSTELLUNG,10)) = CDate(TD.ADATE) AND BK.BON_NR=TD.BELEGNR SET TD.BON_ID=BK.BON_ID")
            
            '2. Gutscheine ermitteln
            lblProgress.Caption = "2. Gutscheine ermitteln ..."
            lblProgress.Refresh
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN GUTZ G ON K.ADATE =G.ADATE AND K.AZEIT=G.AZEIT AND K.KASNUM=G.KASNUM AND K.BELEGNR=G.BELEGNR SET K.GUTSCHEIN_NR = G.GUTSCHNR")
            
            '3. Storno ermitteln
            '   wenn menge>0                     (keine Storno)
            '   wenn menge<0  mit TSE Signatur   (keine Storno)
            '   wenn menge<0 ohne TSE Signatur   (Storno)
            lblProgress.Caption = "3. Storno schreiben ..."
            lblProgress.Refresh
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN hatTSE bit")
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN KASSBON KB ON K.ADATE =KB.DATUM AND K.AZEIT=KB.UHRZEIT AND K.KASNUM=KB.KASNUM AND K.BELEGNR=KB.BONNR SET K.hatTSE = true WHERE KB.TSESTART IS NOT NULL")
            gdBase.Execute ("UPDATE ToDsFinvK SET P_STORNO = '1' WHERE MENGE < 0 AND hatTSE=false")
            
            '4. schreib Warengruppe Text
            lblProgress.Caption = "4. Warengruppe Text wird geschrieben  ..."
            lblProgress.Refresh
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN AGNDBF F ON K.AGN =F.AGN SET K.WARENGR = F.AGTEXT")
              
            '5.Bonpos Tabelle
            lblProgress.Caption = "5. Bonpos Tabelle wird erstellt ..."
            lblProgress.Refresh
            gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,Format (ADATE & ' ' & AZEIT, 'yyyy-mm-dd\Thh:nn:ss') as Z_ERSTELLUNG,Z_NR,BELEGNR,BON_ID,POS_ZEILE,GUTSCHEIN_NR,BEZEICH as ARTIKELTEXT,POS_TERMINAL_ID,GV_TYP,GV_NAME,INHAUS,P_STORNO,AGENTUR_ID,ARTNR as ART_NR,EAN as GTIN,AGN as WARENGR_ID,WARENGR,MENGE,FAKTOR,EINHEIT,VKPR as STK_BR into Bonpos FROM ToDsFinvK ")
            
            '**************************************************************************************> Bonpos_USt.csv
                                                    lblDatei.Caption = "Bonpos_USt.csv"
                                                    lblDatei.Refresh
  
            '1.MWST nach Jahr ermitteln
            lblProgress.Caption = "1.MWST wird ermitteln ..."
            lblProgress.Refresh
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN UST_SCHLUESSEL NUMBER,POS_UST NUMBER")
            
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN MWSTSATZ M ON K.ADATE>=M.vonD AND K.ADATE<=M.bisD SET K.UST_SCHLUESSEL=M.id,K.POS_UST=M.VOLL WHERE K.MWST='V'")
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN MWSTSATZ M ON K.ADATE>=M.vonD AND K.ADATE<=M.bisD SET K.UST_SCHLUESSEL=M.id,K.POS_UST=M.ERM  WHERE K.MWST='E'")
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN MWSTSATZ M ON K.ADATE>=M.vonD AND K.ADATE<=M.bisD SET K.UST_SCHLUESSEL=M.id,K.POS_UST=M.OHNE WHERE K.MWST='O'")
            
            'vielleicht ist [bisD] leer/Null
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN MWSTSATZ M ON K.ADATE>=M.vonD SET K.UST_SCHLUESSEL=M.id,K.POS_UST=M.VOLL WHERE K.MWST='V' AND M.bisD is null")
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN MWSTSATZ M ON K.ADATE>=M.vonD SET K.UST_SCHLUESSEL=M.id,K.POS_UST=M.ERM  WHERE K.MWST='E' AND M.bisD is null")
            gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN MWSTSATZ M ON K.ADATE>=M.vonD SET K.UST_SCHLUESSEL=M.id,K.POS_UST=M.OHNE WHERE K.MWST='O' AND M.bisD is null")
            
            gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,Format (ADATE & ' ' & AZEIT, 'yyyy-mm-dd\Thh:nn:ss') as Z_ERSTELLUNG,Z_NR,BON_ID,POS_ZEILE,UST_SCHLUESSEL,VKPR as POS_BRUTTO,'' as POS_NETTO,POS_UST into Bonpos_USt FROM ToDsFinvK")
                                                                                                                                    
            '2. NettoPreis rechnen
             lblProgress.Caption = "2. NettoPreis wird gerechnet ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE Bonpos_USt SET POS_NETTO = (POS_BRUTTO/(100+POS_UST))*100")
             gdBase.Execute ("UPDATE Bonpos_USt SET POS_NETTO=FORMAT(POS_NETTO,'0.0000') , POS_BRUTTO=FORMAT(POS_BRUTTO,'0.0000')")
             gdBase.Execute ("UPDATE Bonpos_USt SET POS_UST = POS_BRUTTO-POS_NETTO")
             gdBase.Execute ("UPDATE Bonpos_USt SET POS_UST=FORMAT(POS_UST,'0.0000')")
            '**************************************************************************************> Bonpos_Preisfindung.csv
                                                    lblDatei.Caption = "Bonpos_Preisfindung.csv"
                                                    lblDatei.Refresh
              
            
             '1. RabattBetrag (discount) rechnen ( 0 < Preis < VKPR  ,dann es gibt Rabatt )
             lblProgress.Caption = "1. RabattAnteil wird gerechnet  ..."
             lblProgress.Refresh
             
             gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN TYP VARCHAR(20)")
             gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN PF_BRUTTO NUMBER")
             gdBase.Execute ("UPDATE ToDsFinvK SET TYP = 'base_amount' WHERE VKPR = Preis")
             gdBase.Execute ("UPDATE ToDsFinvK SET TYP = 'discount' WHERE Preis>0 AND Preis<VKPR")
             
             gdBase.Execute ("UPDATE ToDsFinvK SET PF_BRUTTO= VKPR WHERE TYP = 'base_amount'")
             gdBase.Execute ("UPDATE ToDsFinvK SET PF_BRUTTO= VKPR - Preis WHERE TYP = 'discount'")
             gdBase.Execute ("UPDATE ToDsFinvK SET PF_BRUTTO= LEFT(PF_BRUTTO,6)")
              
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,Format (ADATE & ' ' & AZEIT, 'yyyy-mm-dd\Thh:nn:ss') as Z_ERSTELLUNG,Z_NR,BON_ID,POS_ZEILE,TYP,UST_SCHLUESSEL,PF_BRUTTO,'' as PF_NETTO, '' as PF_UST INTO Bonpos_Preisfindung FROM ToDsFinvK")
   
                
            '**************************************************************************************> Bonpos_Zusatzinfo.csv
                                                    lblDatei.Caption = "Bonpos_Zusatzinfo.csv"
                                                    lblDatei.Refresh
             '1. grunde Spalten abfragen"
             lblProgress.Caption = "1. die entsprechenden Spalten für Bonpos_Zusatzinfo werden ermittelt ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,Format (ADATE & ' ' & AZEIT, 'yyyy-mm-dd\Thh:nn:ss') as Z_ERSTELLUNG,Z_NR,BON_ID,POS_ZEILE,'' as ZI_ART_NR,'' as ZI_GTIN,'' as ZI_NAME,'' as ZI_WARENGR_ID,'' as ZI_WARENGR,'' as ZI_MENGE,'' as ZI_FAKTOR,'' as ZI_EINHEIT,'' as ZI_UST_SCHLUESSEL,'' as ZI_BASISPREIS_BRUTTO,'' as ZI_BASISPREIS_NETTO,'' as ZI_BASISPREIS_UST INTO Bonpos_Zusatzinfo FROM ToDsFinvK")
                   
            
             '**************************************************************************************> Bonkopf_USt.csv
                                                    lblDatei.Caption = "Bonkopf_USt.csv"
                                                    lblDatei.Refresh
             If NewTableSuchenDB("Bonkopf_USt_tmp", gdBase) Then
               sSQL = "drop Table Bonkopf_USt_tmp"
               gdBase.Execute sSQL, dbFailOnError
             End If
             '1. hier werden in allgemein die Summe-Netto,Summe-Brutto,Summe-MWST jedes Kassenbons gerechnet
             lblProgress.Caption = "1. Summe-Netto,Summe-Brutto,Summe-MWST jedes Kassenbons werden gerechnet ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID,Z_ERSTELLUNG,Z_NR,BON_ID,UST_SCHLUESSEL,POS_BRUTTO,POS_NETTO,POS_UST INTO Bonkopf_USt_tmp FROM Bonpos_USt")
             gdBase.Execute ("SELECT Z_KASSE_ID,Z_ERSTELLUNG,Z_NR,BON_ID,UST_SCHLUESSEL,SUM(POS_BRUTTO)as BON_BRUTTO,SUM(POS_NETTO)as BON_NETTO,SUM(POS_UST)as BON_UST INTO Bonkopf_USt FROM Bonkopf_USt_tmp group by Z_KASSE_ID,Z_ERSTELLUNG,Z_NR,BON_ID,UST_SCHLUESSEL")
             gdBase.Execute ("DROP TABLE Bonkopf_USt_tmp")
             gdBase.Execute ("UPDATE Bonkopf_USt SET BON_BRUTTO=FORMAT(BON_BRUTTO,'0.0000')")
             
             
             '**************************************************************************************> Bonkopf_AbrKreis.csv
                                                    lblDatei.Caption = "Bonkopf_AbrKreis.csv"
                                                    lblDatei.Refresh
             lblProgress.Caption = ""
             lblProgress.Refresh
             '1. hier bleibt die Spalte [ ABRECHNUNGSKREIS] Leer (in Winkiss nicht unterstützt [siehe Dokumentation von Bonkopf_AbrKreis.csv])
             gdBase.Execute ("SELECT Z_KASSE_ID,Z_ERSTELLUNG,Z_NR,BON_ID,'' as ABRECHNUNGSKREIS INTO Bonkopf_AbrKreis FROM Bonkopf")
                   
                   
             '**************************************************************************************> Bonkopf_Zahlarten.csv
                                                    lblDatei.Caption = "Bonkopf_Zahlarten.csv"
                                                    lblDatei.Refresh
              
             '1. Zahlarten ermitteln
             lblProgress.Caption = "1. Zahlart jedes Kassenbons wird ermitteltt ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID,Z_ERSTELLUNG,Z_NR,BON_NR,BON_ID,'' as ZAHLART_TYP,'' as ZAHLART_NAME,'' as ZAHLWAEH_CODE,'' as ZAHLWAEH_BETRAG,UMS_BRUTTO as BASISWAEH_BETRAG INTO Bonkopf_Zahlarten FROM Bonkopf")
             
             lblProgress.Caption = "2. Zahlarten werden geschrieben ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE Bonkopf_Zahlarten BZA INNER JOIN KASSBON K ON BZA.Z_KASSE_ID=K.KASNUM AND BZA.Z_ERSTELLUNG= Format (K.DATUM & ' ' & K.UHRZEIT, 'yyyy-mm-dd\Thh:nn:ss') AND BZA.BON_NR=K.BONNR SET BZA.ZAHLART_TYP=K.KK_ART")
             gdBase.Execute ("UPDATE Bonkopf_Zahlarten SET ZAHLART_TYP='Unbar' WHERE ZAHLART_TYP<>'BA' AND ZAHLART_TYP<>'GZ'")
             gdBase.Execute ("UPDATE Bonkopf_Zahlarten SET ZAHLART_TYP='Bar' WHERE ZAHLART_TYP='BA'")
             gdBase.Execute ("UPDATE Bonkopf_Zahlarten SET ZAHLART_TYP='Unbar',ZAHLART_NAME='Gemischte Zahlung' WHERE ZAHLART_TYP='GZ'")
             
                   
             '**************************************************************************************> Bon_Referenzen.csv
                                                    lblDatei.Caption = "Bon_Referenzen.csv"
                                                    lblDatei.Refresh
             lblProgress.Caption = "Bon_Referenzen wird mit den Daten gefüllt ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID,Z_ERSTELLUNG,Z_NR,BON_ID,'' as POS_ZEILE,'' as REF_TYP,'' as REF_NAME,'' as REF_DATUM,'' as REF_Z_KASSE_ID,'' as REF_Z_NR,'' as REF_BON_ID INTO Bon_Referenzen FROM Bonkopf")
               
                   
                   
             '**************************************************************************************> TSE_Transaktionen.csv
                                                    lblDatei.Caption = "TSE_Transaktionen.csv"
                                                    lblDatei.Refresh
             lblProgress.Caption = "von Bon_Kopf into TSE_Transaktionen"
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID,Z_ERSTELLUNG,Z_NR,BON_ID,BON_NR,'' as TSE_ID,'' as TSE_TANR,'' as TSE_TA_START,'' as TSE_TA_ENDE,'1' as TSE_TA_VORGANGSART,'' as TSE_TA_SIGZ,'' as TSE_TA_SIG,'' as TSE_TA_FEHLER,'' as TSE_VORGANSDATEN INTO TSE_Transaktionen FROM Bonkopf")
             lblProgress.Caption = "TSE Daten werden geschrieben ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE TSE_Transaktionen TS INNER JOIN KASSBON KB ON TS.Z_KASSE_ID=KB.KASNUM AND TS.Z_ERSTELLUNG = Format (KB.DATUM & ' ' & KB.UHRZEIT, 'yyyy-mm-dd\Thh:nn:ss') AND TS.BON_NR=KB.BONNR SET TS.TSE_ID = KB.TSEID,TS.TSE_TANR = KB.TSETRANSACTION,TS.TSE_TA_START = KB.TSESTART,TS.TSE_TA_ENDE = KB.TSEEND,TS.TSE_TA_SIGZ = KB.FINISHSIGZAHLER,TS.TSE_TA_SIG = KB.TSEFINISHSIG,TS.TSE_TA_FEHLER = KB.TSEFEHLER")
             gdBase.Execute ("UPDATE TSE_Transaktionen SET TSE_TA_VORGANGSART='' WHERE LEN(TSE_TA_FEHLER) > 0")
             gdBase.Execute ("UPDATE TSE_Transaktionen SET TSE_TA_VORGANGSART='' WHERE (TSE_TA_FEHLER='' OR TSE_TA_FEHLER is null) AND (TSE_TANR is null OR TSE_TANR='')")
                                  
                                  
             '**************************************************************************************> Stamm_Abschluss.csv
                                                    lblDatei.Caption = "Stamm_Abschluss.csv"
                                                    lblDatei.Refresh
             '1.START/ENDE BON_ID
             gdBase.Execute ("SELECT Z_KASSE_ID,CDate(LEFT(Z_ERSTELLUNG,10))as Z_ERSTELLUNG1,BON_ID,ZAHLART_TYP,BASISWAEH_BETRAG into tmp_Stamm_Abschluss FROM Bonkopf_Zahlarten")
             lblProgress.Caption = "1.Start/Ende von Bon_ID werden geschrieben ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID ,Z_ERSTELLUNG1 ,min(BON_ID)as Z_START_ID ,max(BON_ID)as Z_ENDE_ID INTO tmp_Stamm_START_ID_ENDE_ID FROM tmp_Stamm_Abschluss  group by  Z_KASSE_ID,Z_ERSTELLUNG1")
             
             '2.Unbar Umsätze
             lblProgress.Caption = "2.Unbar Umsätze werden summiert ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID ,Z_ERSTELLUNG1, sum(BASISWAEH_BETRAG)as Z_SE_ZAHLUNGEN  INTO tmp_Stamm_SummeUnbar FROM tmp_Stamm_Abschluss WHERE ZAHLART_TYP='Unbar'  group by  Z_KASSE_ID ,Z_ERSTELLUNG1")
             
             '3.Bar Umsätze
             lblProgress.Caption = "3.Bar Umsätze werden summiert ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID ,Z_ERSTELLUNG1, sum(BASISWAEH_BETRAG)as Z_SE_BARZAHLUNGEN  INTO tmp_Stamm_SummeBar FROM tmp_Stamm_Abschluss WHERE ZAHLART_TYP='Bar'  group by  Z_KASSE_ID ,Z_ERSTELLUNG1")
             
             '4.Bar/Unbar Umsätze schreiben
             lblProgress.Caption = "4.Bar/Unbar Umsätze werden geschrieben ..."
             lblProgress.Refresh
             gdBase.Execute ("Alter table  tmp_Stamm_START_ID_ENDE_ID add COLUMN Z_SE_ZAHLUNGEN NUMBER ,Z_SE_BARZAHLUNGEN NUMBER")
             gdBase.Execute ("UPDATE tmp_Stamm_START_ID_ENDE_ID SET Z_SE_ZAHLUNGEN=0,Z_SE_BARZAHLUNGEN=0")
             gdBase.Execute ("UPDATE tmp_Stamm_START_ID_ENDE_ID SIDEID INNER JOIN tmp_Stamm_SummeBar SB ON SIDEID.Z_KASSE_ID=SB.Z_KASSE_ID AND SIDEID.Z_ERSTELLUNG1=SB.Z_ERSTELLUNG1 SET SIDEID.Z_SE_BARZAHLUNGEN=SB.Z_SE_BARZAHLUNGEN")
             gdBase.Execute ("UPDATE tmp_Stamm_START_ID_ENDE_ID SIDEID INNER JOIN tmp_Stamm_SummeUnbar SUN ON SIDEID.Z_KASSE_ID=SUN.Z_KASSE_ID AND SIDEID.Z_ERSTELLUNG1=SUN.Z_ERSTELLUNG1 SET SIDEID.Z_SE_ZAHLUNGEN=SUN.Z_SE_ZAHLUNGEN")
             
             '5.Firma Daten schreiben
             lblProgress.Caption = "5.Firma Daten werden geschrieben ..."
             lblProgress.Refresh
             gdBase.Execute ("Alter table tmp_Stamm_START_ID_ENDE_ID add COLUMN NAME varchar(50), STRASSE varchar(50), PLZ varchar(7) , ORT varchar(50) ,LAND varchar(12), STEUERNR varchar(35),USTID NUMBER")
             gdBase.Execute ("UPDATE tmp_Stamm_START_ID_ENDE_ID SET LAND='Deutschland'")
             gdBase.Execute ("UPDATE tmp_Stamm_START_ID_ENDE_ID tmp1 , FIRMA F SET tmp1.NAME = F.NAME , tmp1.STRASSE=F.STRASSE , tmp1.PLZ = F.PLZ , tmp1.ORT = F.ORT , tmp1.STEUERNR = F.STEUERNR")
             gdBase.Execute ("UPDATE tmp_Stamm_START_ID_ENDE_ID tmp1 INNER JOIN MWSTSATZ M ON tmp1.Z_ERSTELLUNG1>=M.vonD AND tmp1.Z_ERSTELLUNG1<=M.bisD SET tmp1.USTID=M.id")
             'vielleicht ist [bisD] in der Tabelle MWSTSATZ = Null
             gdBase.Execute ("UPDATE tmp_Stamm_START_ID_ENDE_ID tmp1 INNER JOIN MWSTSATZ M ON tmp1.Z_ERSTELLUNG1>=M.vonD SET tmp1.USTID=M.id WHERE M.bisD is null")
             
             '6.End-Tabelle [Stamm_Abschluss] erstellen
             lblProgress.Caption = "6.Tabelle [Stamm_Abschluss] erstellen ..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT Z_KASSE_ID,Z_ERSTELLUNG1,'' as Z_NR,'' as BUCHUNGSTAG,'' as TAXONOMIE_VERSION,Z_START_ID,Z_ENDE_ID,Z_SE_ZAHLUNGEN,Z_SE_BARZAHLUNGEN,NAME,STRASSE,PLZ,ORT,LAND,STEUERNR,USTID INTO Stamm_Abschluss FROM tmp_Stamm_START_ID_ENDE_ID")
             gdBase.Execute ("Alter table Stamm_Abschluss ADD Column Z_ERSTELLUNG varchar(25)")
             gdBase.Execute ("UPDATE Stamm_Abschluss SET Z_ERSTELLUNG=Z_ERSTELLUNG1 & ' 00:00:00'")
             gdBase.Execute ("UPDATE Stamm_Abschluss SET Z_ERSTELLUNG=Format(Z_ERSTELLUNG,'yyyy-mm-dd\Thh:nn:ss')")
             gdBase.Execute ("Alter table Stamm_Abschluss drop Column Z_ERSTELLUNG1")
             
             '7.DROP die tmp Tabellen von Stamm_Abschluss
             lblProgress.Caption = "7.Temp-Tabellen werden entfernt ..."
             lblProgress.Refresh
             gdBase.Execute ("DROP Table tmp_Stamm_Abschluss")
             gdBase.Execute ("DROP Table tmp_Stamm_START_ID_ENDE_ID")
             gdBase.Execute ("DROP Table tmp_Stamm_SummeBar")
             gdBase.Execute ("DROP Table tmp_Stamm_SummeUnbar")
             
             
             
             '**************************************************************************************> Stamm_Orte.csv
                                                    lblDatei.Caption = "Stamm_Orte.csv"
                                                    lblDatei.Refresh
             lblProgress.Caption = ""
             lblProgress.Refresh
             gdBase.Execute ("SELECT '' as Z_KASSE_ID,'' as Z_ERSTELLUNG,'' as Z_NR,'' as LOC_NAME,'' as LOC_STRASSE,'' as LOC_PLZ,'' as LOC_ORT,'' as LOC_LAND,'' as LOC_USTID INTO Stamm_Orte")
             
             '**************************************************************************************> Stamm_Kassen.csv
                                                    lblDatei.Caption = "Stamm_Kassen.csv"
                                                    lblDatei.Refresh
             lblProgress.Caption = ""
             lblProgress.Refresh
             gdBase.Execute ("SELECT '' as Z_KASSE_ID,'' as Z_ERSTELLUNG,'' as Z_NR,'' as KASSE_BRAND,'' as KASSE_MODELL,'' as KASSE_SERIENNR,'' as KASSE_SW_BRAND,'' as KASSE_SW_VERSION,'' as KASSE_BASISWAEH_CODE,'' as KEINE_UST_ZUORDNUNG INTO Stamm_Kassen")
             
                          
             
             '**************************************************************************************> Stamm_Terminals.csv
                                                    lblDatei.Caption = "Stamm_Terminals.csv"
                                                    lblDatei.Refresh
             lblProgress.Caption = ""
             lblProgress.Refresh
             gdBase.Execute ("SELECT '' as Z_KASSE_ID,'' as Z_ERSTELLUNG,'' as Z_NR,'' as TERMINAL_ID,'' as TERMINAL_BRAND,'' as TERMINAL_MODEL,'' as TERMINAL_SERIENNR,'' as TERMINAL_SW_BRAND,'' as TERMINAL_SW_VERSION INTO Stamm_Terminals")
             
             
             '**************************************************************************************> Stamm_Agenturen.csv
                                                    lblDatei.Caption = "Stamm_Agenturen.csv"
                                                    lblDatei.Refresh
             lblProgress.Caption = ""
             lblProgress.Refresh
             gdBase.Execute ("SELECT '' as Z_KASSE_ID,'' as Z_ERSTELLUNG,'' as Z_NR,'' as AGENTUR_ID,'' as AGENTUR_NAME,'' as AGENTUR_STRASSE,'' as AGENTUR_PLZ,'' as AGENTUR_ORT,'' as AGENTUR_LAND,'' as AGENTUR_STNR,'' as AGENTUR_USTID INTO Stamm_Agenturen")
                          
             
             '**************************************************************************************> Stamm_USt.csv
                                                    lblDatei.Caption = "Stamm_USt.csv"
                                                    lblDatei.Refresh
             'NULL in MWSTSATZ auf 01.01.2100 setzen (besser zum Datum-Vergleichen)
             gdBase.Execute ("update MWSTSATZ set bisD='01.01.2100' where bisD is null")
             
             '1.get KassenAbschlusse von den Tabellen AFCSTATP,KASSJOUR (nach der selektierten Frist)
             lblProgress.Caption = "1.AFCSTATP,KASSJOUR (KassAbschluss) werden abgefragt ..."
             lblProgress.Refresh
             gdBase.Execute ("select AF.ADATE,AF.KASNUM,AF.BELEGNR,KJ.MWST,'' as UST_SCHLUESSEL,'' as UST_SATZ,'' as UST_BESCHR into tmp_Stamm_USt FROM AFCSTATP AF,KASSJOUR KJ where AF.ADATE=KJ.ADATE and AF.KASNUM=KJ.KASNUM and AF.BELEGNR=KJ.BELEGNR and Datevalue(AF.ADATE) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "')")
             
             '2.schreibe UST_SCHLUESSEL,UST_SATZ mit Verknüpfung mit der Tabelle [MWSTSATZ]
             lblProgress.Caption = "2.UST_SCHLUESSEL,UST_SATZ werden geschrieben ..."
             lblProgress.Refresh
             gdBase.Execute ("update tmp_Stamm_USt tmp , MWSTSATZ mw set tmp.UST_SCHLUESSEL=mw.id , tmp.UST_SATZ=mw.VOLL WHERE tmp.MWST='V' AND CDate(tmp.ADATE) between mw.vonD and mw.bisD")
             gdBase.Execute ("update tmp_Stamm_USt tmp , MWSTSATZ mw set tmp.UST_SCHLUESSEL=mw.id , tmp.UST_SATZ=mw.ERM WHERE  tmp.MWST='E' AND CDate(tmp.ADATE) between mw.vonD and mw.bisD")
             gdBase.Execute ("update tmp_Stamm_USt tmp , MWSTSATZ mw set tmp.UST_SCHLUESSEL=mw.id , tmp.UST_SATZ=mw.OHNE WHERE tmp.MWST='O' AND CDate(tmp.ADATE) between mw.vonD and mw.bisD")
             
            'gdBase.Execute ("UPDATE tmp_Stamm_USt SET ADATE=ADATE & ' 00:00:00'")
            'gdBase.Execute ("UPDATE tmp_Stamm_USt SET ADATE=Format(ADATE,'yyyy-mm-dd\Thh:nn:ss')")
 
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,ADATE as Z_ERSTELLUNG,BELEGNR as Z_NR,UST_SCHLUESSEL,UST_SATZ,UST_BESCHR INTO Stamm_USt FROM tmp_Stamm_USt")
             gdBase.Execute ("DROP TABLE tmp_Stamm_USt")
             
             lblProgress.Caption = ""
             lblProgress.Refresh
             gdBase.Execute ("update MWSTSATZ set bisD=null where bisD=CDate('01.01.2100')")
              
             '**************************************************************************************> Stamm_TSE.csv
                                                    lblDatei.Caption = "Stamm_TSE.csv"
                                                    lblDatei.Refresh
             
             gdBase.Execute ("SELECT ADATE,KASNUM,'' as Z_NR,'' as TSE_ID,'' as TSE_SERIAL,'' as TSE_SIG_ALGO,'generalizedTimeWithMilliseconds' as TSE_ZEITFORMAT,'UTF-8' as TSE_PD_ENCODING,'' as TSE_PUBLIC_KEY,'' as TSE_ZERTIFIKAT_I,'' as TSE_ZERTIFIKAT_II INTO tmp_Stamm_TSE_1 FROM AFCSTATP WHERE DateValue(ADATE) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "')")
             gdBase.Execute ("SELECT DISTINCT DATUM,KASNUM,TSEID INTO tmp_Stamm_TSE_2 FROM KASSBON WHERE DateValue(DATUM) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "') AND TSEID is not null")
             gdBase.Execute ("UPDATE tmp_Stamm_TSE_1 tmp1 INNER JOIN tmp_Stamm_TSE_2 tmp2 ON tmp1.ADATE = tmp2.DATUM AND tmp1.KASNUM = tmp2.KASNUM SET tmp1.TSE_ID = tmp2.TSEID")
            
             gdBase.Execute ("UPDATE tmp_Stamm_TSE_1 SET TSE_ID='' WHERE TSE_ID IS NULL")
             
             gdBase.Execute ("UPDATE tmp_Stamm_TSE_1 tmp1 INNER JOIN TSEStorageInfo SI ON CStr(tmp1.TSE_ID) = CStr(SI.TSEID) SET tmp1.TSE_SERIAL = SI.SerialNum, tmp1.TSE_SIG_ALGO = SI.SignaturAlg,tmp1.TSE_PUBLIC_KEY = SI.PublicKey")
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID, ADATE as Z_ERSTELLUNG,Z_NR,TSE_ID,TSE_SERIAL,TSE_SIG_ALGO,TSE_ZEITFORMAT,TSE_PD_ENCODING,TSE_PUBLIC_KEY,TSE_ZERTIFIKAT_I,TSE_ZERTIFIKAT_II INTO Stamm_TSE FROM tmp_Stamm_TSE_1")
             
             gdBase.Execute ("DROP TABLE tmp_Stamm_TSE_1")
             gdBase.Execute ("DROP TABLE tmp_Stamm_TSE_2")
             
             '**************************************************************************************> Z_GV_Typ.csv
                                                    lblDatei.Caption = "Z_GV_Typ.csv"
                                                    lblDatei.Refresh
             
             gdBase.Execute ("update MWSTSATZ set bisD='01.01.2100' where bisD is null")
             
             
             '1.grundlegende Spalten aus der Tabelle KASSJOUR ermitteln
             lblProgress.Caption = "1.grundlegende Spalten aus KASSJOUR ermitteln"
             lblProgress.Refresh
             gdBase.Execute ("SELECT KASNUM,ADATE,BELEGNR,MWST,SUM(VKPR) as ZBrutto INTO tmp_Z_GV_Typ FROM KASSJOUR WHERE DateValue(ADATE) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "') group by KASNUM,ADATE,BELEGNR,MWST")
             
             '2.UST_SCHLUESSEL schreiben
             lblProgress.Caption = "2.UST_SCHLUESSEL wird geschrieben ..."
             lblProgress.Refresh
             gdBase.Execute ("ALTER TABLE tmp_Z_GV_Typ ADD COLUMN MWST_WERT NUMBER")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ tmp , MWSTSATZ M SET tmp.MWST_WERT=M.VOLL WHERE tmp.ADATE between M.vonD AND M.bisD AND tmp.MWST='V' ")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ tmp , MWSTSATZ M SET tmp.MWST_WERT=M.ERM WHERE tmp.ADATE between M.vonD AND M.bisD AND tmp.MWST='E' ")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ tmp , MWSTSATZ M SET tmp.MWST_WERT=M.OHNE WHERE tmp.ADATE between M.vonD AND M.bisD AND tmp.MWST='O' ")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ set MWST='1' WHERE MWST='V'")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ set MWST='2' WHERE MWST='E'")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ set MWST='6' WHERE MWST='O'")
             
             
             '3.Z_UMS_NETTO rechnen
             lblProgress.Caption = "3.Z_UMS_NETTO wird gerechnet..."
             lblProgress.Refresh
             gdBase.Execute ("ALTER TABLE tmp_Z_GV_Typ ADD COLUMN ZNetto NUMBER")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ SET ZNetto=ZBrutto/(1+(MWST_WERT/100))")
             
             '4.Z_UST rechnen
             lblProgress.Caption = "4.Z_UST wird gerechnet..."
             lblProgress.Refresh
             gdBase.Execute ("ALTER TABLE tmp_Z_GV_Typ ADD COLUMN ZUst NUMBER")
             gdBase.Execute ("UPDATE tmp_Z_GV_Typ SET ZUst=ZBrutto-ZNetto")
                

             '5.Tabelle [Z_GV_Typ] generieren
             lblProgress.Caption = "5.Tabelle [Z_GV_Typ] wird generiert..."
             lblProgress.Refresh
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,ADATE as Z_ERSTELLUNG,BELEGNR as Z_NR,'Umsatz' as GV_TYP,'' as GV_NAME,'0' as AGENTUR_ID,MWST as UST_SCHLUESSEL,ZBrutto as Z_UMS_BRUTTO,ZNetto as Z_UMS_NETTO,ZUst as Z_UST INTO Z_GV_Typ FROM tmp_Z_GV_Typ")
             gdBase.Execute ("DROP TABLE tmp_Z_GV_Typ")
             
             '6.Tabelle [Z_GV_Typ] formatieren
             lblProgress.Caption = "6.Tabelle [Z_GV_Typ] wird formatiert..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE Z_GV_Typ SET Z_UMS_BRUTTO=FORMAT(Z_UMS_BRUTTO,'0.00') , Z_UMS_NETTO=FORMAT(Z_UMS_NETTO,'0.00'), Z_UST=FORMAT(Z_UST,'0.00')")
             
             
             gdBase.Execute ("UPDATE MWSTSATZ set bisD=null where bisD=CDate('01.01.2100')")
              
              
              
              
              
               
              
              
              
             lblProgress.Caption = "FERTIG"
             lblProgress.Refresh
             
              
             
             'CSVhelper.exe aufrufen  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
            
                Dim tmpDB_Pfad As String
                Dim tmpDB_Pass As String
                Dim autoOeffnen As String
                
                tmpDB_Pfad = gcDBPfad & "\kissdata.mdb"
                tmpDB_Pass = "Kiss2005"
                'den Ordner der erstellten Dateien automatisch Öffnen
                 If AutomatischDateiOffnen Then
                   autoOeffnen = "ja"
                  Else
                   autoOeffnen = "nein"
                 End If
                
                Shell App.Path & "\" & "CSVhelper.exe " & tmpDB_Pfad & " " & tmpDB_Pass & " " & gbDsFinvkPfad & " " & autoOeffnen, vbNormalFocus
                
             'CSVhelper.exe aufrufen  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
             
             
             
             Command2.Enabled = True
             WirdVerarbeitet = True
             
             
             Exit Sub
            
              
     
        '///////////////////// Fertig //////////////////////
       

Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StarteExportieren"
    Fehler.gsFehlertext = "Im Programmteil DsFinvK Expo. ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
 
  
Private Sub Form_Activate()
 WirdVerarbeitet = True
End Sub

Private Sub Form_Load()
 
  AutomatischDateiOffnen = False
  
  Me.BackColor = glH1
  
  lblDatei.BackColor = glH1
  lblProgress.BackColor = glH1
  ChkOeffnen.BackColor = glH1
  Label1.BackColor = glH1
  
  lblDatei.ForeColor = vbYellow
  lblProgress.ForeColor = vbYellow
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
 gdBase.Execute ("update MWSTSATZ set bisD=null where bisD=CDate('01.01.2100')")
                    
End Sub


Private Sub alteTabellenVonDsFinvKLoschen()
On Error GoTo LOKAL_ERROR
 
  If NewTableSuchenDB("Bonpos", gdBase) Then
    sSQL = "drop Table Bonpos"
    gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Bonpos_USt", gdBase) Then
     sSQL = "drop Table Bonpos_USt"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Bonpos_Preisfindung", gdBase) Then
     sSQL = "drop Table Bonpos_Preisfindung"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Bonpos_Zusatzinfo", gdBase) Then
     sSQL = "drop Table Bonpos_Zusatzinfo"
     gdBase.Execute sSQL, dbFailOnError
  End If
   
   
  If NewTableSuchenDB("Bonkopf", gdBase) Then
     sSQL = "drop Table Bonkopf"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Bonkopf_USt", gdBase) Then
     sSQL = "drop Table Bonkopf_USt"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
   If NewTableSuchenDB("Bonkopf_AbrKreis", gdBase) Then
     sSQL = "drop Table Bonkopf_AbrKreis"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
 If NewTableSuchenDB("Bonkopf_Zahlarten", gdBase) Then
     sSQL = "drop Table Bonkopf_Zahlarten"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Bon_Referenzen", gdBase) Then
     sSQL = "drop Table Bon_Referenzen"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  
   If NewTableSuchenDB("Bon_Referenzen", gdBase) Then
     sSQL = "drop Table Bon_Referenzen"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
 If NewTableSuchenDB("TSE_Transaktionen", gdBase) Then
     sSQL = "drop Table TSE_Transaktionen"
     gdBase.Execute sSQL, dbFailOnError
  End If
 
 
 
 ''''''''''''''''''''''''''''''''''''''''''''
 
 
 
  If NewTableSuchenDB("tmp_Stamm_Abschluss", gdBase) Then
     sSQL = "drop Table tmp_Stamm_Abschluss"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("tmp_Stamm_START_ID_ENDE_ID", gdBase) Then
     sSQL = "drop Table tmp_Stamm_START_ID_ENDE_ID"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  
   If NewTableSuchenDB("tmp_Stamm_SummeUnbar", gdBase) Then
     sSQL = "drop Table tmp_Stamm_SummeUnbar"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
 If NewTableSuchenDB("tmp_Stamm_SummeBar", gdBase) Then
     sSQL = "drop Table tmp_Stamm_SummeBar"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Stamm_Abschluss", gdBase) Then
     sSQL = "drop Table Stamm_Abschluss"
     gdBase.Execute sSQL, dbFailOnError
  End If
   
  If NewTableSuchenDB("Stamm_Orte", gdBase) Then
     sSQL = "drop Table Stamm_Orte"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
   If NewTableSuchenDB("Stamm_Kassen", gdBase) Then
     sSQL = "drop Table Stamm_Kassen"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
   If NewTableSuchenDB("Stamm_Terminals", gdBase) Then
     sSQL = "drop Table Stamm_Terminals"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
   If NewTableSuchenDB("Stamm_Agenturen", gdBase) Then
     sSQL = "drop Table Stamm_Agenturen"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("tmp_Stamm_USt", gdBase) Then
     sSQL = "drop Table tmp_Stamm_USt"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Stamm_USt", gdBase) Then
     sSQL = "drop Table Stamm_USt"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("tmp_Stamm_TSE_1", gdBase) Then
     sSQL = "drop Table tmp_Stamm_TSE_1"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("tmp_Stamm_TSE_2", gdBase) Then
     sSQL = "drop Table tmp_Stamm_TSE_2"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Stamm_TSE", gdBase) Then
     sSQL = "drop Table Stamm_TSE"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("tmp_Z_GV_Typ", gdBase) Then
     sSQL = "drop Table tmp_Z_GV_Typ"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  If NewTableSuchenDB("Z_GV_Typ", gdBase) Then
     sSQL = "drop Table Z_GV_Typ"
     gdBase.Execute sSQL, dbFailOnError
  End If
  
  
Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "alteTabellenVonDsFinvKLoschen"
    Fehler.gsFehlertext = "Im Programmteil DsFinvK Expo. ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub


