VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ExportFormular 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   13995
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
      Format          =   110952449
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
      Format          =   110952449
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

Dim AutomatischDateiÖffnen As Boolean
Dim rsRes As Recordset
Dim iDsF As Integer
Dim DateiZuErstellen As String
     

 

 

'Private Sub bisJahr_Change()
'
'bisJahr.BackColor = vbWhite
'
' Dim textval As String
'
' textval = Trim(bisJahr.Text)
' textval = Replace(textval, ".", "")
' textval = Replace(textval, ",", "")
'
'  If IsNumeric(textval) Then
'      bisJahr.Text = CStr(textval)
'    Else
'      bisJahr.Text = ""
'
'  End If
'
'End Sub
'
'Private Sub bisJahr_Click()
'bisJahr.BackColor = vbWhite
'End Sub

Private Sub ChkOeffnen_Click()
 
 
 If ChkOeffnen.value = vbChecked Then
    
          AutomatischDateiÖffnen = True
     Else
          AutomatischDateiÖffnen = False
    
 End If
 
 
 
End Sub

Private Sub Command1_Click()

ChooseFile.Top = Me.Top - 200
ChooseFile.Left = Me.Left + Me.Width / 4
ChooseFile.Show 1

End Sub

Private Sub Command2_Click()
On Error GoTo LOKAL_ERROR

 
  'DatePart("ww", Now())

If Trim(gbDsFinvkPfad) = "" Then

 MsgBox ("Bitte erstmal Pfad wählen ! ! !")
 
' ElseIf vonJahr.Text = "von" Or Trim(vonJahr.Text) = "" Then
'  MsgBox ("bitte Datum auswählen")
'  vonJahr.BackColor = vbRed
' ElseIf bisJahr.Text = "bis" Or Trim(bisJahr.Text) = "" Then
'  MsgBox ("bitte Datum auswählen")
'  bisJahr.BackColor = vbRed
 Else
   
 Command2.Enabled = False
 alteTabellenVonDsFinvKLoschen
 StarteExportieren
 
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
            gdBase.Execute ("UPDATE ToDsFinvK TD INNER JOIN Bonkopf BK ON BK.Z_KASSE_ID=TD.KASNUM AND BK.Z_ERSTELLUNG=Format (TD.ADATE & ' ' & TD.AZEIT, 'yyyy-mm-dd\Thh:nn:ss') AND BK.BON_NR=TD.BELEGNR SET TD.BON_ID=BK.BON_ID")
            
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
            gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,Format (ADATE & ' ' & AZEIT, 'yyyy-mm-dd\Thh:nn:ss') as Z_ERSTELLUNG,Z_NR,BON_ID,POS_ZEILE,GUTSCHEIN_NR,BEZEICH as ARTIKELTEXT,POS_TERMINAL_ID,GV_TYP,GV_NAME,INHAUS,P_STORNO,AGENTUR_ID,ARTNR as ART_NR,EAN as GTIN,AGN as WARENGR_ID,WARENGR,MENGE,FAKTOR,EINHEIT,VKPR as STK_BR into Bonpos FROM ToDsFinvK ")
            
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
             
             gdBase.Execute ("UPDATE tmp_Stamm_USt SET ADATE=ADATE & ' 00:00:00'")
             gdBase.Execute ("UPDATE tmp_Stamm_USt SET ADATE=Format(ADATE,'yyyy-mm-dd\Thh:nn:ss')")
             
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID,ADATE as Z_ERSTELLUNG,BELEGNR as Z_NR,UST_SCHLUESSEL,UST_SATZ,UST_BESCHR INTO Stamm_USt FROM tmp_Stamm_USt")
             gdBase.Execute ("DROP TABLE tmp_Stamm_USt")
             
             lblProgress.Caption = ""
             lblProgress.Refresh
             gdBase.Execute ("update MWSTSATZ set bisD=null where bisD=CDate('01.01.2100')")
              
             '**************************************************************************************> Stamm_TSE.csv
                                                    lblDatei.Caption = "Stamm_TSE.csv"
                                                    lblDatei.Refresh
             
             gdBase.Execute ("SELECT ADATE,KASNUM,'' as Z_NR,'' as TSE_ID,'' as TSE_SERIAL,'' as TSE_SIG_ALGO,'generalizedTimeWithMilliseconds' as TSE_ZEITFORMAT,'UTF-8' as TSE_PD_ENCODING,'' as TSE_PUBLIC_KEY,'' as TSE_ZERTIFIKAT_I,'' as TSE_ZERTIFIKAT_II INTO tmp_Stamm_TSE_1 FROM AFCSTATP WHERE DateValue(ADATE) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "')")
             gdBase.Execute ("SELECT DISTINCT DATUM,KASNUM,TSEID INTO tmp_Stamm_TSE_2 FROM KASSBON WHERE DateValue(DATUM) between CDate('" & DTPicker1.value & "') and CDate('" & DTPicker2.value & "')")
             gdBase.Execute ("UPDATE tmp_Stamm_TSE_1 tmp1 INNER JOIN tmp_Stamm_TSE_2 tmp2 ON tmp1.ADATE = tmp2.DATUM AND tmp1.KASNUM = tmp2.KASNUM SET tmp1.TSE_ID = tmp2.TSEID")
            
             gdBase.Execute ("UPDATE tmp_Stamm_TSE_1 SET TSE_ID='' WHERE TSE_ID IS NULL")
             
             gdBase.Execute ("UPDATE tmp_Stamm_TSE_1 tmp1 INNER JOIN TSEStorageInfo SI ON CStr(tmp1.TSE_ID) = CStr(SI.TSEID) SET tmp1.TSE_SERIAL = SI.SerialNum, tmp1.TSE_SIG_ALGO = SI.SignaturAlg,tmp1.TSE_PUBLIC_KEY = SI.PublicKey")
             gdBase.Execute ("SELECT KASNUM as Z_KASSE_ID, ADATE as Z_ERSTELLUNG,Z_NR,TSE_ID,TSE_SERIAL,TSE_SIG_ALGO,TSE_ZEITFORMAT,TSE_PD_ENCODING,TSE_PUBLIC_KEY,TSE_ZERTIFIKAT_I,TSE_ZERTIFIKAT_II INTO Stamm_TSE FROM tmp_Stamm_TSE_1")
             
             gdBase.Execute ("DROP TABLE tmp_Stamm_TSE_1")
             gdBase.Execute ("DROP TABLE tmp_Stamm_TSE_2")
             
             lblProgress.Caption = "FERTIG"
             lblProgress.Refresh
             Exit Sub
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
                                                                                            
            '2. extra Columns für die Prozessierung hinzufügen
            lblProgress.Caption = "2. extra Columns für die Prozessierung hinzufügen  ..."
            lblProgress.Refresh
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN MWST_WERT NUMBER")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN NettoPreis NUMBER")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Ust_Schluessel NUMBER")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Warengruppe TEXT")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Einheit TEXT")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Stueck_Preis NUMBER")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN RabattAnteil NUMBER")
            gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Storno TEXT")
                
                
            '3. MWST als Prozent % überschreiben(hier als Prozent %, aber in Shritt 12 als Wert)
             lblProgress.Caption = "3. MWST mit dem Wert ersetzen  ..."
             lblProgress.Refresh
               '3.1 get die Prozent-Wert aus der Tabelle MWSTSATZ  ( VOLL , Ermäßig , Ohne )
               Dim rsMw As Recordset
               Dim wertVoll As Double
               Dim wertErm As Double
               Dim wertOhne As Double
               
               Set rsMw = gdBase.OpenRecordset("SELECT * FROM MWSTSATZ")
                If Not IsNull(rsMw) Then
                
                    If Not rsMw.EOF Then
                    
                     rsMw.MoveFirst
                     wertVoll = IIf(rsMw!VOLL = "", 0, rsMw!VOLL)
                     wertErm = IIf(rsMw!ERM = "", 0, rsMw!ERM)
                     wertOhne = IIf(rsMw!OHNE = "", 0, rsMw!OHNE)
                    
                    End If
                
                End If
                rsMw.Close
                Set rsMw = Nothing
                
                '3.2 MWST in der Tabelle ToDsFinvK mit den in Schritt 3.1 gebrachten Daten ersetzen
                gdBase.Execute ("UPDATE ToDsFinvK set MWST_WERT = " & wertVoll & " WHERE MWST = 'V'")
                gdBase.Execute ("UPDATE ToDsFinvK set MWST_WERT = " & wertErm & " WHERE MWST = 'E'")
                gdBase.Execute ("UPDATE ToDsFinvK set MWST_WERT = " & wertOhne & " WHERE MWST = 'O'")
                
             '4. NettoPreis rechnen
             lblProgress.Caption = "4. NettoPreis wird gerechnet ..."
             lblProgress.Refresh
               
             gdBase.Execute ("UPDATE ToDsFinvK SET NettoPreis = (VKPR/(100+MWST_WERT))*100")
             
             '5. Ust_Schluessel (ID für MWST) schreiben
             lblProgress.Caption = "5. Ust_Schluessel wird geschrieben ..."
             lblProgress.Refresh
                
             gdBase.Execute ("UPDATE ToDsFinvK SET Ust_Schluessel = 3 WHERE MWST_WERT = 10.7")
             gdBase.Execute ("UPDATE ToDsFinvK SET Ust_Schluessel = 4 WHERE MWST_WERT = 5.5")
             gdBase.Execute ("UPDATE ToDsFinvK SET Ust_Schluessel = 5 WHERE MWST_WERT = 0")
             gdBase.Execute ("UPDATE ToDsFinvK SET Ust_Schluessel = 11 WHERE MWST_WERT = 19")
             gdBase.Execute ("UPDATE ToDsFinvK SET Ust_Schluessel = 12 WHERE MWST_WERT = 7")
             gdBase.Execute ("UPDATE ToDsFinvK SET Ust_Schluessel = 21 WHERE MWST_WERT = 16")
             gdBase.Execute ("UPDATE ToDsFinvK SET Ust_Schluessel = 22 WHERE MWST_WERT = 5")
                
             '6. Storno als J / N schreiben ( Hinweis: Storno heißt Menge < 0 )
             lblProgress.Caption = "6. Storno als J / N schreiben ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE ToDsFinvK SET Storno = 'J' WHERE Menge < 0")
             gdBase.Execute ("UPDATE ToDsFinvK SET Storno = 'N' WHERE Menge > 0")
                
                
             '7. schreib Warengruppe Text
             lblProgress.Caption = "7. Warengruppe Text wird geschrieben  ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE ToDsFinvK K INNER JOIN AGNDBF F ON K.AGN =F.AGN SET K.Warengruppe = F.AGTEXT")
                             
                             
             '8. Einheit ist immer = Stueck
             lblProgress.Caption = "8. Einheit ist immer Stueck  ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE ToDsFinvK SET Einheit = 'Stueck'")
             
             '9. StueckPreis ist immer gleich VKPR
             lblProgress.Caption = "9. StueckPreis ist immer gleich VKPR  ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE ToDsFinvK SET Stueck_Preis = VKPR")
                     
             '10. RabattAnteil rechnen ( 0 < Preis < VKPR  ,dann es gibt Rabatt )
             lblProgress.Caption = "10. RabattAnteil wird gerechnet  ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE ToDsFinvK SET RabattAnteil = Round(((VKPR-Preis)/VKPR)*100) WHERE Preis>0 AND Preis<VKPR")
                
             '11. NettoPreis,Stueck_Preis formatieren
             lblProgress.Caption = "11. NettoPreis,Stueck_Preis werden formatiert  ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE ToDsFinvK SET NettoPreis=FORMAT(NettoPreis,'0.00') , Stueck_Preis=FORMAT(Stueck_Preis,'0.00')")
             
             
             '12. MWST als Wert rechnen [ BruttoPreis - NettoPreis ], wobei der BruttoPreis der VKPR ist (siehe oben Hinweis im Schritt 1)
             lblProgress.Caption = "12. csv wird erstellt und mit Ergebnisse gefüllt  ..."
             lblProgress.Refresh
             gdBase.Execute ("UPDATE ToDsFinvK SET MWST_WERT = VKPR - NettoPreis")
 
             '13. Bonpos.csv erstellen und mit Ergebnisse füllen
             lblProgress.Caption = "13. Bonpos.csv wird erstellt und mit Ergebnisse gefüllt  ..."
             lblProgress.Refresh
         
             csvDateiErstellen "Einzelaufzeichnungsmodul\Bonpos", "Bonpos.csv"
        
 
                    
        
               '**************************************************************************************> Bonkopf.csv
                                                    lblDatei.Caption = "Bonkopf.csv"
                                                    lblDatei.Refresh
        
        
               'Hinweis: die letzte Tabelle [ ToDsFinvK ] und ihre Spalten dadrin (nicht Alle) werden hier auch benutzt.
               ''''''''' Außerdem werden noch neue Spalten in [ ToDsFinvK ] für die Prozessierung hinzugefügt
            
                
                '14. extra Columns für die Prozessierung hinzufügen
                lblProgress.Caption = "14. extra Columns für die Prozessierung hinzufügen  ..."
                lblProgress.Refresh
                gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Kundnr NUMBER")
                gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Bediener_ID NUMBER")
                gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Bediener_Name TEXT")
                gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Kunde_Name TEXT")
                gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Kunde_Strasse TEXT")
                gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Kunde_PLZ TEXT")
                
                 
                '15. Kundnr,Bediener von der Tabelle Kassjour in ToDsFinvK kopieren
                lblProgress.Caption = "15. Kundnr,Bediener werden ermittelt  ..."
                lblProgress.Refresh
                gdBase.Execute ("UPDATE ToDsFinvK T INNER JOIN Kassjour J ON T.ADATE=J.ADATE AND T.AZEIT=J.AZEIT AND T.BELEGNR=J.BELEGNR SET T.Kundnr = J.Kundnr , T.Bediener_ID = J.Bediener")
                
                
                '16. Kunden weitere info(Name,Strasse,PLZ) ermitteln
                lblProgress.Caption = "16.  Kunden weitere info(Name,Strasse,PLZ) werden ermittelt  ..."
                lblProgress.Refresh
                gdBase.Execute ("UPDATE ToDsFinvK T INNER JOIN Kunden KU ON T.Kundnr=KU.Kundnr SET T.Kunde_Name = KU.Vorname , T.Kunde_Strasse = KU.Strasse, T.Kunde_PLZ = KU.PLZ")
                    
            
                '17. Bediener weitere Info(Bediener_Name) ermitteln
                lblProgress.Caption = "17. Bediener weitere Info(Bediener_Name) werden ermittelt  ..."
                lblProgress.Refresh
                gdBase.Execute ("UPDATE ToDsFinvK T INNER JOIN BEDNAME B ON T.Bediener_ID=B.BEDNU SET T.Bediener_Name = B.BEDNAME")
                    
                
                '18. Bonkopf.csv erstellen
                lblProgress.Caption = "18. Bonkopf.csv wird erstellt  ..."
                lblProgress.Refresh
                
                csvDateiErstellen "Einzelaufzeichnungsmodul\Bonkopf", "Bonkopf.csv"
                
           
                '**************************************************************************************> Bonkopf_USt.csv
                                                    lblDatei.Caption = "Bonkopf_USt.csv"
                                                    lblDatei.Refresh
                                                    
                 csvDateiErstellen "Einzelaufzeichnungsmodul\Bonkopf", "Bonkopf_USt.csv"
                 
                '**************************************************************************************> Bonkopf_Zahlarten.csv
                                                    lblDatei.Caption = "Bonkopf_Zahlarten.csv"
                                                    lblDatei.Refresh
                                                    
                '19. extra Columns für die Prozessierung hinzufügen
                lblProgress.Caption = "19. extra Columns für die Prozessierung hinzufügen  ..."
                lblProgress.Refresh
                gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN ZahlungArt TEXT")
                
                '20. ZahlungsArt von der Tabelle Kassjour in ToDsFinvK kopieren
                lblProgress.Caption = "20. ZahlungsArt werden ermittelt  ..."
                lblProgress.Refresh
                gdBase.Execute ("UPDATE ToDsFinvK T INNER JOIN Kassjour J ON T.ADATE=J.ADATE AND T.AZEIT=J.AZEIT AND T.BELEGNR=J.BELEGNR SET T.ZahlungArt = J.KK_ART")
                                                   
                '21. Bonkopf_Zahlarten.csv erstellen
                lblProgress.Caption = "21. Bonkopf_Zahlarten.csv wird erstellt  ..."
                lblProgress.Refresh
                
                csvDateiErstellen "Einzelaufzeichnungsmodul\Bonkopf", "Bonkopf_Zahlarten.csv"
                
               '***************************************************************************> Bonkopf_AbrKreis.csv
                                                    lblDatei.Caption = "Bonkopf_AbrKreis.csv"
                                                    lblDatei.Refresh
                                                    
                ' leere Datei weil die Dokumentation der DsFinvK hierbei unbegreiflich bzw. ignoriert war
                csvDateiErstellen "Einzelaufzeichnungsmodul\Bonkopf", "Bonkopf_AbrKreis.csv"
               
               '***************************************************************************> Bonkopf_Referenzen.csv
                                                    lblDatei.Caption = "Bonkopf_Referenzen.csv"
                                                    lblDatei.Refresh
                                                    
                ' leere Datei weil die Dokumentation der DsFinvK hierbei unbegreiflich bzw. ignoriert war
                csvDateiErstellen "Einzelaufzeichnungsmodul\Bonkopf", "Bonkopf_Referenzen.csv"
       
       
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''  START Stammdatenmodul ''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '*********************************************************************************> Stamm_Abschluss.csv
                                                    lblDatei.Caption = "Stamm_Abschluss.csv"
                                                    lblDatei.Refresh
        
        '22. hier wird andere Struktur für die Tabelle[ ToDsFinvK ] benutzt, deswegen wird ToDsFinvK gelöscht und erneut erstellt
        lblProgress.Caption = "22. ToDsFinvK löschen und wieder erstellen  ..."
        lblProgress.Refresh
        gdBase.Execute ("drop Table ToDsFinvK")
       
        '23. die grundlegende Columns aus der Tabelle [AFCSTATP] in der Tabelle [ToDsFinvK] kopieren ( nur für das gegenwärtige jahr )
        lblProgress.Caption = "23. grundlegende Columns werden aus der Tabelle AFCSTATP abgefragt ..."
        lblProgress.Refresh
        gdBase.Execute ("SELECT ADATE,KASNUM,UMS_BAR as Z_SE_BARZAHLUNGEN,(UMS_BAR+UMS_SCHECK+UMS_KRED+UMS_KARTE+ZHLGGUTSCH+UMS_LAST+DUKA)as Z_SE_ZAHLUNGEN INTO ToDsFinvK FROM AFCSTATP WHERE year(ADATE)= year(DATE()) ")
       
        '24. extra Columns für die Prozessierung hinzufügen
        lblProgress.Caption = "24. extra Columns für die Prozessierung werden hinzufügt ..."
        lblProgress.Refresh
        'es lohnt sich, dass TAXONOMIE_VERSION der DsFinvK Version-Nummer ist. bisher ist die Version 2.2
        gdBase.Execute ("ALTER TABLE ToDsFinvK ADD COLUMN Z_START_ID TEXT,Z_ENDE_ID TEXT,TAXONOMIE_VERSION NUMBER")
        gdBase.Execute ("UPDATE ToDsFinvK SET TAXONOMIE_VERSION = 2.2")
        
        '25. Z_SE_BARZAHLUNGEN formatieren
        lblProgress.Caption = "25. Z_SE_BARZAHLUNGEN wird formatiert ..."
        lblProgress.Refresh
        gdBase.Execute ("UPDATE ToDsFinvK SET Z_SE_BARZAHLUNGEN=FORMAT(Z_SE_BARZAHLUNGEN,'0.00')")
       
        '26. Min(BELEGNR) und MAX(BELEGNR) aus der Tabelle KASSJOUR bringen (anhand den Columns [ADATE,KASNUM] )
        '   und in einer neuen Tabelle [ DsFinvKTemp ] schreiben, sodass die in der Tabelle [ ToDsFinvK ] migriert werden
        lblProgress.Caption = "26. Min(BELEGNR) und MAX(BELEGNR) werden ermittelt ..."
        lblProgress.Refresh
        gdBase.Execute ("SELECT ADATE,KASNUM,MIN(BELEGNR) as Z_START_ID,MAX(BELEGNR) as Z_ENDE_ID INTO DsFinvKTemp FROM KASSJOUR WHERE YEAR(ADATE)=YEAR(DATE()) group by ADATE,KASNUM")
        
        '27. die im letzten Schritt ermittelten Min & Max(BELEGNR) in der Tabelle[ ToDsFinvK ]migrieren
        lblProgress.Caption = "27. Min und Max (BELEGNR) werden in der Tabelle migriert ..."
        lblProgress.Refresh
        gdBase.Execute ("UPDATE ToDsFinvK T INNER JOIN DsFinvKTemp TE ON T.ADATE=TE.ADATE AND T.KASNUM=TE.KASNUM SET T.Z_START_ID=TE.Z_START_ID,T.Z_ENDE_ID=TE.Z_ENDE_ID")
        
        '28. die neu in Schritt 5 erstellte Tabelle [ DsFinvKTemp ] löschen, weil die nicht mehr nötig ist
        lblProgress.Caption = "28. die tabelle DsFinvKTemp wird gelöscht ..."
        lblProgress.Refresh
        gdBase.Execute ("drop table DsFinvKTemp")
              
              
        csvDateiErstellen "Stammdatenmodul", "Stamm_Abschluss.csv"
       
        '*********************************************************************************> Stamm_Abschluss.csv
                                                    lblDatei.Caption = "Stamm_Abschluss.csv"
                                                    lblDatei.Refresh
                                                    
        csvDateiErstellen "Stammdatenmodul", "Stamm_Orte.csv"
       
        '*********************************************************************************> Stamm_Kassen.csv
                                                    lblDatei.Caption = "Stamm_Kassen.csv"
                                                    lblDatei.Refresh
                                                    
        csvDateiErstellen "Stammdatenmodul", "Stamm_Kassen.csv"
        
        '*********************************************************************************> Stamm_Terminals.csv
                                                    lblDatei.Caption = "Stamm_Terminals.csv"
                                                    lblDatei.Refresh
                                                    
        csvDateiErstellen "Stammdatenmodul", "Stamm_Terminals.csv"
        
        '*********************************************************************************> Stamm_Agenturen.csv
                                                    lblDatei.Caption = "Stamm_Agenturen.csv"
                                                    lblDatei.Refresh
                                                    
        csvDateiErstellen "Stammdatenmodul", "Stamm_Agenturen.csv"
        
        
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''  START Kassenabschlussmodul ''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '*********************************************************************************> Z_GV_Typ.csv
                                                    lblDatei.Caption = "Z_GV_Typ.csv"
                                                    lblDatei.Refresh
                                                    
        'leere Datei weil die Dokumentation der DsFinvK hierbei unbegreiflich bzw. ignoriert war
        csvDateiErstellen "Kassenabschlussmodul", "Z_GV_Typ.csv"
        
        '*********************************************************************************> Z_Waehrungen.csv
                                                    lblDatei.Caption = "Z_Waehrungen.csv"
                                                    lblDatei.Refresh
                                                    
        'leere Datei weil die Dokumentation der DsFinvK hierbei unbegreiflich bzw. ignoriert war
        csvDateiErstellen "Kassenabschlussmodul", "Z_Waehrungen.csv"
        
        '*********************************************************************************> Z_Zahlart.csv
                                                    lblDatei.Caption = "Z_Zahlart.csv"
                                                    lblDatei.Refresh
                                                    
        '29. hier wird andere Struktur für die Tabelle[ ToDsFinvK ] benutzt, deswegen wird ToDsFinvK gelöscht und erneut erstellt
        lblProgress.Caption = "29. ToDsFinvK löschen und wieder erstellen  ..."
        lblProgress.Refresh
        gdBase.Execute ("drop Table ToDsFinvK")
       
        '30. die grundlegende Columns aus der Tabelle [KASSJOUR] in der Tabelle [ToDsFinvK] kopieren ( nur für das gegenwärtige jahr )
        lblProgress.Caption = "30. grundlegende Columns werden aus der Tabelle KASSJOUR abgefragt ..."
        lblProgress.Refresh
        gdBase.Execute ("SELECT ADATE,KASNUM,KK_ART as ZAHLART_NAME ,SUM(PREIS)as Z_ZAHLART_BETRAG  INTO ToDsFinvK FROM KASSJOUR WHERE YEAR(ADATE)=YEAR(DATE()) group by   ADATE,KASNUM,KK_ART order by ADATE asc")
        
        '31. Z_ZAHLART_BETRAG formatieren
        lblProgress.Caption = "31. NettoPreis,Stueck_Preis werden formatiert  ..."
        lblProgress.Refresh
        gdBase.Execute ("UPDATE ToDsFinvK SET Z_ZAHLART_BETRAG=FORMAT(Z_ZAHLART_BETRAG,'0.00')")
                   
        csvDateiErstellen "Kassenabschlussmodul", "Z_Zahlart.csv"
        
     
        '///////////////////// Fertig //////////////////////
        
        lblProgress.Caption = "Fertig"
        lblProgress.Refresh
        
        'den Ordner der erstellten Dateien oeffnen
         If AutomatischDateiÖffnen Then
          Shell "explorer.exe " & gbDsFinvkPfad, vbMaximizedFocus
         End If
     

Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "StarteExportieren"
    Fehler.gsFehlertext = "Im Programmteil DsFinvK Expo. ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

















' diese Funktion wird für das Erstellen verschiedener .csv-Dateien benutzt, weil jede Datei spezielle Headers und Werte hat
Private Sub csvDateiErstellen(ByVal OrdnerName As String, ByVal DateiName As String)
 On Error GoTo LOKLAL_ERROR
 
 '////////////////////////////////////////////////////////////// START /////////////////////////////////////////
 
 
 
    'die abzufragende Columns
    Dim QueryCmd As String
    
    'Headers zum Schreiben in der .csv-Datei
    Dim tmpHeaders As String
    
    'Wert zum Schreiben in der .csv-Datei
    Dim tmpHeadersWerte As String
 
 
    QueryCmd = ""
    OrdnerName = OrdnerName & "\"
 
 
 
    'Schritt 1.               je nach zu erstellender Datei werden die abzufragene Columns festgelegt
    '************************************************************************************************
    
    Select Case DateiName
        
        Case "Bonpos.csv"
   
                QueryCmd = "SELECT * FROM ToDsFinvK order by ADATE DESC, AZEIT asc"
               
        Case "Bonkopf.csv"
              
                QueryCmd = "SELECT ADATE,AZEIT,BELEGNR,KASNUM,Bediener_ID,Bediener_Name,Kunde_Name,Kunde_Strasse,Kunde_PLZ FROM ToDsFinvK group by ADATE,AZEIT,BELEGNR,KASNUM,Bediener_ID,Bediener_Name,Kunde_Name,Kunde_Strasse,Kunde_PLZ order by ADATE DESC, AZEIT asc"
                 
        Case "Bonkopf_USt.csv"
              
                QueryCmd = "SELECT ADATE,AZEIT,BELEGNR,KASNUM,FORMAT(SUM(NettoPreis),'0.00')as BON_NETTO,FORMAT(SUM(MWST_WERT),'0.00')as BON_UST,FORMAT(SUM(VKPR),'0.00')as BON_BRUTTO FROM ToDsFinvK group by ADATE,AZEIT,BELEGNR,KASNUM order by ADATE DESC, AZEIT asc"
                                  
        Case "Bonkopf_Zahlarten.csv"
              
                QueryCmd = "SELECT ADATE,AZEIT,BELEGNR,KASNUM,ZahlungArt FROM ToDsFinvK group by ADATE,AZEIT,BELEGNR,KASNUM,ZahlungArt order by ADATE DESC, AZEIT asc"
            
        Case "Stamm_Abschluss.csv"
              
                QueryCmd = "SELECT * FROM ToDsFinvK order by ADATE DESC"
                
        Case "Stamm_Orte.csv"
              
                QueryCmd = "SELECT FILIALNAME FROM FILIALEN"
                      
        Case "Z_Zahlart.csv"
              
                QueryCmd = "SELECT * FROM ToDsFinvK"
         
         
    End Select
    
 
 
 
 
     
     'Schritt 2.               die festgelegten Columns abfragen
     '************************************************************************************************
     If Trim(QueryCmd) <> "" Then
        
         Set rsRes = gdBase.OpenRecordset(QueryCmd)
 
        Else
         Set rsRes = Nothing
     End If
 
 
     
   
     'Schritt 3.               die grundlegende DsFinfK-Ordner erstellen ( aber wenn die schon vorhanden sind, dann überschreiben )
     '*****************************************************************************************************************************
      
        Dim BackUpVon_gbDsFinvkPfad As String
            BackUpVon_gbDsFinvkPfad = gbDsFinvkPfad
        
        If Dir$(gbDsFinvkPfad & OrdnerName, vbDirectory) <> "" Then
        
                'Ordner existiert
                 DateiZuErstellen = gbDsFinvkPfad & OrdnerName & DateiName
           Else
           
                'Ordner existiert nicht
                 Dim arr() As String
                 arr = Split(OrdnerName, "\")
                 
                 For i = 0 To UBound(arr) - 1
                 
                  gbDsFinvkPfad = gbDsFinvkPfad & arr(i) & "\"
                   If Dir$(gbDsFinvkPfad, vbDirectory) = "" Then
                    FileSystem.MkDir (gbDsFinvkPfad)
                   End If
                 Next i
                 
                  DateiZuErstellen = gbDsFinvkPfad & DateiName
        End If
       
             
             gbDsFinvkPfad = BackUpVon_gbDsFinvkPfad
     
     
      
    'Schritt 4.               Headers und Werte schreiben
    '*******************************************************
         
    iDsF = FreeFile
    Open DateiZuErstellen For Output As #iDsF
 


                If Not rsRes Is Nothing Then
                
                    Select Case DateiName
                                             

                            
                            Case "Bonpos.csv"
                            
                            
                                     'Schreib Columns Headers
                                      tmpHeaders = "ADATE;AZEIT;BELEGNR;ARTNR;EAN;AGN;BEZEICH;KASNUM;MENGE;PREIS;MWST;NettoPreis;MWST_WERT;BruttoPreis;Ust_Schluessel;Warengruppe;Einheit;Stueck_Preis;RabattAnteil;Storno" & " "
                                      Print #iDsF, tmpHeaders
                                          
                                      Do While Not rsRes.EOF
                                      
                                          tmpHeadersWerte = ""
                                          tmpHeadersWerte = tmpHeadersWerte & rsRes!ADATE & ";" & rsRes!AZEIT & ";" & rsRes!BELEGNR & ";" & rsRes!artnr & ";" & rsRes!EAN & ";" & rsRes!AGN & ";" & rsRes!BEZEICH & ";" & rsRes!KASNUM & ";" & rsRes!Menge & ";" & rsRes!Preis & ";" & rsRes!MWST & ";" & rsRes!nettopreis & ";" & rsRes!MWST_WERT & ";" & rsRes!vkpr & ";" & rsRes!Ust_Schluessel & ";" & rsRes!Warengruppe & ";" & rsRes!Einheit & ";" & rsRes!Stueck_Preis & ";" & rsRes!RabattAnteil & ";" & rsRes!Storno & " "
                                          Print #iDsF, tmpHeadersWerte
                                          rsRes.MoveNext
                                      Loop
                                   
                            
                            Case "Bonkopf.csv"
                            
                                     'Schreib Columns Headers
                                      tmpHeaders = "ADATE;AZEIT;BELEGNR;KASNUM;Bediener_ID;Bediener_Name;Kund_Name;Kund_Strasse;Kund_Strasse;Kund_PLZ" & " "
                                      Print #iDsF, tmpHeaders
                                      
                                      Do While Not rsRes.EOF
                                      
                                          tmpHeadersWerte = ""
                                          tmpHeadersWerte = tmpHeadersWerte & rsRes!ADATE & ";" & rsRes!AZEIT & ";" & rsRes!BELEGNR & ";" & rsRes!KASNUM & ";" & rsRes!Bediener_ID & ";" & rsRes!Bediener_Name & ";" & rsRes!Kunde_Name & ";" & rsRes!Kunde_Strasse & ";" & rsRes!Kunde_PLZ & " "
                                          Print #iDsF, tmpHeadersWerte
                                          rsRes.MoveNext
                                      Loop
                                      
                            Case "Bonkopf_USt.csv"
                            
                                     'Schreib Columns Headers
                                      tmpHeaders = "ADATE;AZEIT;BELEGNR;KASNUM;BON_NETTO;BON_UST;BON_BRUTTO" & " "
                                      Print #iDsF, tmpHeaders
                                      
                                      Do While Not rsRes.EOF
                                      
                                          tmpHeadersWerte = ""
                                          tmpHeadersWerte = tmpHeadersWerte & rsRes!ADATE & ";" & rsRes!AZEIT & ";" & rsRes!BELEGNR & ";" & rsRes!KASNUM & ";" & rsRes!BON_NETTO & ";" & rsRes!BON_UST & ";" & rsRes!BON_BRUTTO & " "
                                          Print #iDsF, tmpHeadersWerte
                                          rsRes.MoveNext
                                      Loop
                                      
                                
                            Case "Bonkopf_Zahlarten.csv"
                            
                                     'Schreib Columns Headers
                                      tmpHeaders = "ADATE;AZEIT;BELEGNR;KASNUM;ZahlungArt" & " "
                                      Print #iDsF, tmpHeaders
                                      
                                      Do While Not rsRes.EOF
                                      
                                          tmpHeadersWerte = ""
                                          tmpHeadersWerte = tmpHeadersWerte & rsRes!ADATE & ";" & rsRes!AZEIT & ";" & rsRes!BELEGNR & ";" & rsRes!KASNUM & ";" & rsRes!ZahlungArt & " "
                                          Print #iDsF, tmpHeadersWerte
                                          rsRes.MoveNext
                                      Loop
                                      
                            Case "Stamm_Abschluss.csv"
                            
                                     'Firma Info abfragen
                                      Dim rsFirma As Recordset
                                      Set rsFirma = gdBase.OpenRecordset("SELECT NAME , STRASSE , PLZ , ORT , Steuernr FROM FIRMA")
                                      If Not rsFirma.EOF Then
                                       
                                       Print #iDsF, "NAME :" & ";" & rsFirma!name & " "
                                       Print #iDsF, "STRASSE :" & ";" & rsFirma!strasse & " "
                                       Print #iDsF, "PLZ :" & ";" & rsFirma!Plz & " "
                                       Print #iDsF, "ORT :" & ";" & rsFirma!Ort & " "
                                       Print #iDsF, "Steuernr :" & ";" & rsFirma!Steuernr & " " & " "
                                       Print #iDsF, ";" & " "
                                       Print #iDsF, ";" & " "
                                       
                                       rsFirma.Close
                                       Set rsFirma = Nothing
                                      End If
                                    
                                     'Schreib Columns Headers
                                      tmpHeaders = "ADATE;KASNUM;Z_SE_BARZAHLUNGEN;TAXONOMIE_VERSION;Z_SE_ZAHLUNGEN;Z_START_ID;Z_ENDE_ID" & " "
                                      Print #iDsF, tmpHeaders
                                      
                                      Do While Not rsRes.EOF
                                      
                                          tmpHeadersWerte = ""
                                          tmpHeadersWerte = tmpHeadersWerte & rsRes!ADATE & ";" & rsRes!KASNUM & ";" & rsRes!Z_SE_BARZAHLUNGEN & ";" & rsRes!TAXONOMIE_VERSION & ";" & rsRes!Z_SE_ZAHLUNGEN & ";" & rsRes!Z_START_ID & ";" & rsRes!Z_ENDE_ID & " "
                                          Print #iDsF, tmpHeadersWerte
                                          rsRes.MoveNext
                                      Loop
                                      
                            Case "Stamm_Orte.csv"
                            
                                     'Schreib Columns Headers
                                      tmpHeaders = "LOC_NAME" & " "
                                      Print #iDsF, tmpHeaders
                                      
                                      Do While Not rsRes.EOF
                                      
                                          tmpHeadersWerte = ""
                                          tmpHeadersWerte = tmpHeadersWerte & rsRes!Filialname & " "
                                          Print #iDsF, tmpHeadersWerte
                                          rsRes.MoveNext
                                      Loop
                                      
                            Case "Z_Zahlart.csv"
                            
                                     'Schreib Columns Headers
                                      tmpHeaders = "ADATE;KASNUM;ZAHLART_NAME;Z_ZAHLART_BETRAG" & " "
                                      Print #iDsF, tmpHeaders
                                      
                                      Do While Not rsRes.EOF
                                      
                                          tmpHeadersWerte = ""
                                          tmpHeadersWerte = tmpHeadersWerte & rsRes!ADATE & ";" & rsRes!KASNUM & ";" & rsRes!ZAHLART_NAME & ";" & rsRes!Z_ZAHLART_BETRAG & " "
                                          Print #iDsF, tmpHeadersWerte
                                          rsRes.MoveNext
                                      Loop
                                      
                                     
                    End Select
                     
                End If
                
                

            If Not rsRes Is Nothing Then
                rsRes.Close
                Set rsRes = Nothing
            End If
            
            
            

    Close #iDsF

 
    
    
    '////////////////////////////////////////////////////////////// ENDE /////////////////////////////////////////
     
     Exit Sub
     
LOKLAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "csvDateiErstellen"
    Fehler.gsFehlertext = "Im Programmteil DsFinvK Expo. ist ein Fehler aufgetreten."

    Fehlermeldung1
 
End Sub


Private Sub Form_Load()

 
  AutomatischDateiÖffnen = False
  
  Me.BackColor = glH1
  
  lblDatei.BackColor = glH1
  lblProgress.BackColor = glH1
  ChkOeffnen.BackColor = glH1
  Label1.BackColor = glH1
  
  lblDatei.ForeColor = vbYellow
  lblProgress.ForeColor = vbYellow
   
'  Dim i As Integer
'  i = Year(Date)
'
'  Do While i >= 2020
'   vonJahr.AddItem (CStr(i))
'   bisJahr.AddItem (CStr(i))
'   i = i - 1
'  Loop
  


End Sub

 

Private Sub Form_Unload(Cancel As Integer)
 
 gdBase.Execute ("update MWSTSATZ set bisD=null where bisD=CDate('01.01.2100')")
                    
End Sub

'Private Sub vonJahr_Change()
'vonJahr.BackColor = vbWhite
'
' Dim textval As String
'
' textval = Trim(vonJahr.Text)
' textval = Replace(textval, ".", "")
' textval = Replace(textval, ",", "")
'
'  If IsNumeric(textval) Then
'      vonJahr.Text = CStr(textval)
'    Else
'      vonJahr.Text = ""
'
'  End If
'End Sub

'Private Sub vonJahr_Click()
' vonJahr.BackColor = vbWhite
'End Sub


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
  
  
Exit Sub

LOKAL_ERROR:
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "alteTabellenVonDsFinvKLoschen"
    Fehler.gsFehlertext = "Im Programmteil DsFinvK Expo. ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub


