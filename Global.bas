Attribute VB_Name = "Global"
Option Explicit

Global Const glpVers                As Long = 3138
Global gbSQLSERVER                  As Boolean

Global gbDsFinvkPfad As String
Global gbQRFlag As Boolean
Global gbBudniNeuesFtpVerfahren As Boolean

Global geretteteF_DateiErfolgreichAbgeschickt As Boolean

Global Const gbDEMO As Boolean = False
Global Const gbKostenlos As Boolean = False

Global byAnzahlSpalten              As Byte
Global sSpaltenname()               As String
Global sSpaltenbez()                As String
Global sSpaltenAli()                As String
Global aBreite()                    As Integer

Global sGlobAenderGRUND             As String

Global byAnzahlSpaltenEX            As Byte
Global sFremdSpalteEX()             As String
Global sKissSpalteEX()              As String

Global gsVorEinPLZ1                 As String
Global gsVorEinPLZ2                 As String

Global giKeinDel                    As Integer

Global gsVEDES_HOST                 As String
Global gsVEDES_USER                 As String
Global gsVEDES_PW                   As String

Global gsVEDES_HOST_DSL             As String
Global gsVEDES_USER_DSL             As String
Global gsVEDES_PW_DSL               As String




Global gbTSE_SCHREIBEN              As Boolean
Global gsTSE_APIKEY                 As String
Global gsTSE_APISECRET              As String
Global gsTSE_TSEID                  As String
Global gsTSE_CLIENTID               As String









Global sStornoText()                    As String
Global sStornoTextT1()                  As String
Global dZollEr(10)                      As Double

Global gbBudni_Bestellung_erfolgreich   As Boolean

Global Const GBZeitschlossVersion   As Boolean = False
Global gsZeitschlossdate            As Date
Global gsZeitPass                   As String
Global bAbschlussjetzt              As Boolean
Global gsNeuerAbschluß              As String

Global gdKartenschwellenwert        As Double
Global gsPfadBestandlive            As String
Global gsJUGENDSCHUTZFARBE          As String
Global gsUnbekanntStrichMail        As String
Global gsNachtVerarbeitungMail      As String
Global gsDabaNachtStart             As String

Global gsWWZeichen                  As String
Global gsWWwert                     As String
Global gsWWSchwellenwert            As String
Global gsWWArt                      As String
Global gbWWKundBi                   As Boolean
Global gsWWBonusArtnr               As String
Global gsWWBonusGDAUER              As String

Global gcArrSerienNr()              As String
Global gcArrBemerk()                As String
Global gcArrArtNr()                 As String

Global gsTextVor                    As String
Global gsTextNach                   As String
Global giBonusNr                    As Integer

Global gbZweitMoni                  As Boolean
Global gbZweitMoniMinimieren        As Boolean
Global gbSonderPreisDarstellen      As Boolean
Global gbArtEindeut                 As Boolean
Global gbPLZGEBIET                  As Boolean
Global gbMitKundeWahlHinweis        As Boolean
Global gbPLZGEBIET_AuchBeiKUWAHL    As Boolean
Global gbCoupon                     As Boolean
Global gbGuStattBar                 As Boolean
Global gbHenkel                     As Boolean
Global gbArtikelTextSuche           As Boolean
Global gbBestandsgrund              As Boolean
Global gcUmleittxt                  As String
Global gcBedKUNEU                   As String
Global giUmleitgrund                As Integer
Global sErrDabapfad                 As String
Global srechnertab                  As String
Global theBigFehler                 As Boolean
Global theBigFTPFehler              As Boolean
Global theBigFTPFehlerTemp          As Boolean
Global theBigFTPFehlerZähler        As Integer
Global giCopyMod                    As Integer
Global gsHelpstring                 As String
Global gsArtNot                     As String
Global gsArtNotBez                  As String
Global gbNachKBbeiEC                As Boolean
Global gbOhnebestProt               As Boolean
Global gbKeineBestVerWarengru       As Boolean
Global gbBestDateien                As Boolean
Global gbUmsAnz                     As Boolean
Global gbBarAnz                     As Boolean
Global gbEinfacheZollErstattung     As Boolean
Global gbEDITKASSNR                 As Boolean
Global gbKBmBI                      As Boolean
Global gbmGDetails                  As Boolean
Global gbArtrabhalten               As Boolean
Global gbArtrabhaltenLiebling       As Boolean

Global gbBEDLEER                    As Boolean
Global gbBONNEIN                    As Boolean
Global gbBONWAHL                    As Boolean
Global gbGTBON                      As Boolean
Global gbNewArt                 As Boolean
Global gbNewArtNrVorschlag      As Boolean
Global gbPAEBON                 As Boolean
Global gbkassgefuehrt           As Boolean
Global gbBestinZ                As Boolean
Global gbBestAkt                As Boolean
Global gbHauptg                 As Boolean
Global gsLagerFTPBox            As String
Global gbKONTIN                 As Boolean
Global gsGesRabforNetto         As String
Global gcWKDBPfad               As String
Global gsARTNR                  As String
Global gsGRUPPENNR              As String
Global gsQMBez                  As String
Global gsQMPreis                As String
Global gdWechselgeld            As Double
Global gdKassenGeldGezählt      As Double
Global gbGutschOverBar          As Boolean
Global gsSEK                    As String
Global gbKDEXM                  As Boolean
Global gsiDisPause              As Single
Global gsBackcolor              As String
Global gsForecolor              As String
Global gsArtikelFarbe           As String
Global gsKundenfarbbeschreib    As String
Global gsKundenFarbe            As String
Global gcSuch                   As String
Global gsAnforderung            As String
Global gsPdfDatei               As String
Global gsLokalTabellen(0 To 6)  As String
Global gcEAN                    As String
Global gdDBPAUSE                As Double
Global gbmv                     As Boolean
Global gsMORGENTEXT             As String
Global gsMITTAGTEXT             As String
Global gsABENDTEXT              As String
Global gsABOPLUS_KARTE          As String
Global gdABOPLUS_WERT           As Double
Global gbPenner_faerben         As Boolean
Global gsPLZ                    As String

Global gdUmsatzproKundeDurchschnitt As Double
Global gdUmsatzMittelproKunde   As Double
Global gdKaufvorgänge           As Double
Global gbMitMwstAnteile         As Boolean

Global gcTerm_Datum             As String
Global gbTerm_Name              As Boolean
Global gbTerm_InfoDauerh        As Boolean
Global gbTerm_BedKass           As Boolean
Global gcTerm_Bed               As String

Global gdEuroproBonKU As Double
Global glAnzvkKU As Long
Global gdumsgesKU As Double

Global gdEuroproBonKU365 As Double
Global glAnzvkKU365 As Long
Global gdumsgesKU365 As Double

Global gBYTENum(0 To 100) As Long
Global gBYTENumLIN(0 To 100) As Long

Global Kasstime                 As String
Global lKasstimeEnde            As Long
Global lKasstimeBegin           As Long

Global iErrseconds              As Integer

Global d68Summe                 As Single
Global d68Nochoffen             As Single
Global c68Kdnr                  As String
Global gbBackaus68              As Boolean
Global gbBackaus20g             As Boolean
Global iWelchekarte             As Integer
Global gLGutschnum              As Long
Global gb118bestell             As Boolean
Global gdRückgeldaus68          As Double
Global gbCCfromBestlief         As Boolean

'für MDDETAILS
Global MBDETAILBVO              As Double
Global MBDETAILVON              As Long
Global MBDETAILBIS              As Long
Global MBDETAILMON              As Long

Type Email
    SenderName As String
    ReplyTo As String
    SenderEMail As String
    CC As String
    BCC As String
    Recipient As String
    SMTPAUTH As Boolean
    SSL As Boolean
    ServerName As String
    ServerPort As Long
    Username As String
    Password As String
    Subject As String
    Message As String
    AutoZIP As Boolean
    Attachment1 As String
    Attachment2 As String
    Attachment3 As String
    Attachment4 As String
    Attachment5 As String
    Attachment6 As String
End Type

Type PauseA
    Pausenkrit As Date
    Pausenlaenge As Date
End Type

Type IndexKombi
    Tabe As String
    Inde As String
    IndeLis As String
End Type

Type Preislage
    PreisVon As Double
    PreisBis As Double
    Preislagentext As String * 30
    PreislagenNr As Byte
End Type

Type ArtikelTyp
    artnr           As Long
    BEZEICH         As String * 35
    AGN             As Integer
    lekpr           As Single
    REKPR           As Single
    vkpr            As Single
    MWST            As String * 1
    linr            As Long
    LIBESNR         As String * 13
    EAN             As String * 13
    EAN2            As String * 13
    EAN3            As String * 13
    ETIMERK         As String * 1
    MOPREIS         As Single
    RKZ             As String * 1
    LPZ             As Integer
    NOTIZEN         As String * 25
    BESTAND         As Long
    VKMENGE         As Long
    VKDATUM         As Date
    MINMEN          As Integer
    INHALT          As Single
    INHALTBEZ       As String * 3
    GRUNDPREIS      As String * 1
    MINBEST         As Integer
    RABATT_OK       As String * 1
    GEFUEHRT        As String * 1
    KVKPR1          As Single
    ekpr            As Single
    PREISSCHU       As String * 1
    BONUS_OK        As String * 1
    UMS_OK          As String * 1
    AWM             As String * 1
    LASTDATE        As Date
    LASTTIME        As String * 10
    AUFDAT          As Date
    EXDAT           As Date
    FARBNR          As Long
    MARKE           As String * 35
    GROESSE         As String * 10
    SPANNE          As Single
    AUFSCHLAG       As Single
    ZubuchMe        As Long
    ZubuchKg        As Single
    OLDBESTAND      As Long
    LAGERPLATZ      As Double
End Type

Type SortErgTyp
    WGN As Double
    WGNTEXT As String * 25
    UMSLJ As Double
    UMSVJ As Double
    DIFFABS As Double
    DIFFPROZ As Double
End Type

Type SortTyp
    AGN As Double
    AGNTEXT As String * 30
End Type

Type UPDATE_
    cDateiName As String
    cDatum As String
    cUhrZeit As String
End Type

Type F2PROMPT_
    cFeld As String
    cEsFeld As String
    cWert As String
    cWert2 As String
    cWahl As String
    bMultiple As Boolean
    cArray(0 To 100) As String
    lLastPos As Long
End Type

Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
    dwFlags As Long
End Type

Type LAYOUT_
    cTabelle As String
    cFeldName As String
    lLfdNr As Long
    cAnzeigen As String
    cBearbeiten As String
    cDBFeld As String
End Type

Type WARENGRU_
    lWgNr As Long
    dArtNr As Double
    cBezeich As String
    cFaktor As String
End Type

Type VERBINDUNG_
    lBaudRate As Long
    sStopBits As Single
    iDatenBits As Integer
    cParitaet As String
    iComPort As Integer
    cSettings As String
    cSchnittstelle As String
End Type

Type REGISTERINFO_
    firma As String
    Plz As String
    Ort As String
    KdWert1 As String
    KdWert2 As String
    KdWert3 As String
    KdWert4 As String
    Confirm1 As String
    Confirm2 As String
    Confirm3 As String
    Confirm4 As String
    Datum As String
End Type

Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

Type MA_STATIST
    lKundZahl As Long
    dNettoVk As Double
    dEkWert As Double
End Type

Type GUTSCHEIN
    gutschnr As Long
    gutschwert As Double
End Type

Type Kunde
    vorname As String
    nachname As String
    Kuerzel As String
    Plz As String
    Ort As String
    strasse As String
    telefon As String
    Mobiltel As String
    Email As String
    titel As String
    firma As String
    anrede As String
    BONUS As String
    LAND As String
    GEBDATUM As String
    KTEXT2 As String
    geschlecht As String
End Type

Type DLG_ZUGRIFF
    lcount As Long
    dZugriff As Double
    dDlg As Double
End Type

Type FIRMA_
    FirmaName As String
    strasse As String
    Plz As String
    Ort As String
    Tel As String
    Fax As String
    BankName As String
    BLZ As String
    Konto As String
    ILN_1 As String
    ILN_2 As String
    Steuernr As String
    BIC As String
    IBAN As String
    FirmaMail As String
End Type

Type ZEITEN_
    WoTag As Integer
    LFDNR As Integer
    Von As String
    Bis As String
    Zeitblock As Integer
End Type

Type ECASHBON
    Kopfzeile_1 As String
    Kopfzeile_2 As String
    Kopfzeile_3 As String
    Kopfzeile_4 As String
    Kopfzeile_5 As String
    Funktion As String
    Kartenart As String
    Betrag As String
    Storno As String
    Datum As String
    Uhrzeit As String
    TerminalID As String
    
    Tracenummer As String
    Belegnummer As String
    TracenummerSTORNO As String
    BLZ As String
    Kontonummer As String
    
    Konto As String
    Karte As String
    Kartenfolgenummer As String
    Verfallsdatum As String
    AIDParameter As String
    Belegduplikat As String
    StornoID As String
    Autorisierungsmerkmal As String
    AutorisierungsNr As String
    Referenzparameter As String
    ReferenzNr As String
    
    VuNr As String
    Fusszeile_1 As String
    Fusszeile_2 As String
    
    ProviderText_01 As String
    ProviderText_02 As String
    ProviderText_03 As String
    ProviderText_04 As String
    TelefonBuchung As String
    BelegNrVon As String
    BelegNrBis As String
    ZahlungAnzahl As String
    Online As String
    Manuell As String
    ZahlungBetrag As String
    ZahlungGAnzahl As String
    ZahlungGBetrag As Single
    ErgebnisText_1 As String
    ErgebnisText_2 As String
    
    ZELVAnz As String
    ZELVGbetrag As Single
    ZPOZecAnz As String
    ZPOZecGbetrag As Single
    ZDinersAnz As String
    ZDinersbetrag As Single
    ZVisaCardAnz As String
    ZVisaCardbetrag As Single
    ZAmericanExpressAnz As String
    ZAmericanExpressbetrag As Single
    ZEuroCardAnz As String
    ZEuroCardbetrag As Single
    
    
    
End Type



Type ECKARTE_
    Original As String
    Datenstrom As String
    ECSpur1 As String
    ECSpur2 As String
    BLZ As String
    BankName As String
    BankOrt As String
    Konto1 As String
    Konto2 As String
    jahr As String
    Monat As String
    KontoInhaber As String
    LastSchriftNr As String
End Type

Type DTA_ASATZ_
    SatzLen As String * 4       'Konstante 0128
    SatzArt As String * 1       'Konstante A
    Hinweis As String * 2       'Gut/Lastschrift Kunde/Bank GK/LK/GB/LB
    BLZ_Empf As String * 8      'BLZ des Geld-Empfängers
    Filler1 As String * 8       'nicht benutzt, mit 0 füllen
    ABSENDER As String * 27     'Name des Disk-Absenders
    Datum As String * 6         'Erstelldatum der Diskette
    Filler2 As String * 4       'nicht benutzt, mit Space füllen
    KontoEmpf As String * 10    'Kontonummer des Geld-Empfängers
    RefNr As String * 10        'Angaben freigestellt
    Reserve1 As String * 15     'nicht benutzt, mit Space füllen
    Erfuellung As String * 8    'Ausführungsdatum (= Erstelldatum)
    Reserve2 As String * 24     'nicht benutzt, mit Space auffüllen
    WaeCode As String * 1       'Space = DM, 1 = Euro
End Type

Type DTA_CSATZ_
    SatzLen As String * 4       'Satzlänge, max. 0622
    SatzArt As String * 1       'Konstante C
    KdBLZ1 As String * 8        'erstbeteiligtes Institut, freigestellt
    KdBLZ2 As String * 8        'endbegünstigtes Institut
    Konto As String * 10        'Geldempfänger, rechtsbündig
    KdNr As String * 13         'interne KdNr, 1.Byte = 0, KdNr mit Nullen, 13.Byte = 0
    TextKey As String * 2       '???
    TextKeyAdd As String * 3    '???
    Filler1 As String * 1       'bankinternes Feld, mit Space auffüllen
    BetragDM As String * 11     'Betrag in DM inkl. Kommastellen ohne Komma, rechtsbündig
    EmpfBLZ As String * 8       'erstbeauftragtes Institut
    EmpfKonto As String * 10    'Auftraggeber, rechtsbündig
    BetragEuro As String * 11   'Betrag in Euro inkl. Kommastellen ohne Komma, rechtsbündig
    Filler2 As String * 3       'nicht benutzt, mit Space auffüllen
    Empfaenger As String * 27   'Überweisungsempfänger, linksbündig, mit Space auffüllen
    Filler3 As String * 8       'nicht benutzt, mit Space auffüllen
    
    AuftragName As String * 27   'Auftraggeber
    Zweck As String * 27        'Verwendungszweck
    WaeCode As String * 1       'Space = DM, 1 = Euro
    Filler4 As String * 2       'nicht benutzt, mit Space auffüllen
    AnzErweit As String * 2     'Anzahl Erweiterungsteile (00 - 15)
    
    KzErwTeil1 As String * 2    'Kennzeichen Erweiterungsteil (Konstante '03')
    Wahltext1 As String * 27    'freier Text im Erweiterungsteil
    KzErwTeil2 As String * 2    'Kennzeichen Erweiterungsteil (Konstante '03')
    Wahltext2 As String * 27    'freier Text im Erweiterungsteil
    Filler5 As String * 11      'nicht benutzt, mit Space auffüllen
    
End Type

Type DTA_ESATZ_
    SatzLen As String * 4       'Satzlänge, Konstante 0128
    SatzArt As String * 1       'Konstante E
    Filler1 As String * 5       'nicht benutzt, mit Space auffüllen
    AnzSatz As String * 7       'Anzahl Sätze von DTA_CSATZ
    SumTotalDM As String * 13   'Summe von DTA_CSATZ, Feld BETRAGDM, wenn WAECODE = Space
    SumKonto As String * 17     'Summe von DTA_CSATZ, Feld KONTO
    SumBLZ As String * 17       'Summe von DTA_CSATZ, Feld KDBLZ2
    SumTotalEuro As String * 13 'Summe von DTA_CSATZ, Feld BETRAGEURO, wenn WAECODE = 1
    Filler2 As String * 51      'nicht benutzt, mit Space auffüllen
End Type

Type DTA_BEGLEIT_
    BegleitZettel As String     'Konstante "Begleitzettel"
    BelegloserDTA As String     'Konstante "Belegloser Datenträgeraustausch"
    Sammel As String            'Konstante "Sammel-Einziehungsauftrag an"
    BankName As String          'Name der Bank (aus BLZ Absender)
    VolNr As String             'Vol-Nr der Diskette
    ErstellungsDatum As String  'Datum der Diskette
    AusFuehrDatum As String     'Datum der Ausführung
    AnzSatzC As String          'Anzahl der Sätze in C
    SummeDM As String           'DM-Summe der Datensätze in C
    SummeEuro As String         'Euro-SUmme der Datensätze in C
    SummeKonto As String        'Kontrollsumme Kontonummern
    SummeBLZ As String          'Kontrollsumme BLZ
    AbsName As String           'Name des Absenders
    AbsBLZ As String            'BLZ, Konto des Absenders
    AbsKonto As String          'Konto des Absenders
    EmpfName As String          'Name des Absenders
    EmpfBLZ As String           'BLZ, Konto des Absenders
    EmpfKonto As String         'Konto des Absenders
    Ort As String               'Ort der Erstellung
    Datum As String             'Datum der Erstellung
    firma As String             'Firma des Erstellers
    Unterschrift As String      'Unterschrift des Erstellers
End Type

Type KASSEINI_
    LFDNR As Integer
    Wert1 As Integer
    WERT2 As Integer
End Type

Type Datensatz
    Kennung As Integer
    name As String * 158
End Type

Global gcBestellEmail   As Email
Global gslfnr           As String
Global iAktAbschlussNr  As Integer
Global gbDabakompfrueh  As Boolean
Global gbDabakompautoNo As Boolean
Global gbOhneAnzeige    As Boolean
Global gbKopOhneAuswertung As Boolean
Global gbWKAUS          As Boolean
'*********************************
'Thomas Globale Variablen   ******
'*********************************
Global gbAwert          As Boolean
Global gbGFKWeek        As Boolean
Global gbGFK            As Boolean
Global gbNIELB          As Boolean
Global gbDate           As Boolean
Global giAkMonat        As Integer
Global gileMonat        As Integer
Global iKasse           As Integer
Global gsKunr           As String
Global giJahr           As Integer
Global gbErstesZeichen  As Boolean
Global gcDateidatum     As String
Global gbfrm27          As Boolean
Global gb159            As Boolean
Global gbFrmComeFrom    As Boolean
Global gfrmComeFrom     As Form
Global gsEinAusBezeich  As String
Global gsfrmComeFrom    As String

Global gsFehlertext     As String
Global gsFormular       As String
Global gsFunktion       As String

Global gbErfolg             As Boolean
Global gbGescheitert        As Boolean

Global giErrorZaehler       As Byte
Global gstab                As String
Global gsZSpalte            As String
Global gsZSpalte1           As String
Global gsZSpalte2           As String
Global gsZSpalte3           As String
Global gsZSpalte4           As String
Global gsZSpalte5           As String
Global iKeypress            As Integer

Global gsStammFTPAdresse    As String
Global gsStammFTPUSER       As String
Global gsStammFTPPASS       As String
Global giStammFTPOFT        As Integer
Global gdateStammlastFTP    As Date

Global gsZenFTPAdresse      As String
Global gsZenFTPUSER         As String
Global gsZenFTPPASS         As String
Global giZenFTPOFT          As Integer
Global gdateZenlastFTP      As Date

Global gbDELBDAT            As Boolean
Global gbDIFFPROT           As Boolean
Global gbUEBERPROT          As Boolean
Global gbBEDKARTE           As Boolean
Global gbQPASS              As Boolean
Global gbDruck27            As Boolean
Global gbETIBEIFARB         As Boolean
Global gbFILMEK             As Boolean

Global gbGeld               As Boolean
Global gbHandelsspanne_Ausblenden As Boolean
Global gbAlterGutschein_Ausblenden As Boolean
Global gdVerBGesrabatt      As Double
Global gdCheckPreis         As Double
Global gbPBARGeld           As Boolean
Global gbMitStaffelPreis    As Boolean
Global gbQZBON              As Boolean
Global gbSterne             As Boolean
Global gbNeukunden          As Boolean
Global gbKASSMBEST          As Boolean
Global gbRESTinBAR          As Boolean
Global gbTPbf               As Boolean
Global gbPark               As Boolean
Global gbDritteArtikelzeile As Boolean
Global gbParknetto          As Boolean
Global gbArtsucheArtFarb    As Boolean
Global gbBILDTAST           As Boolean
Global gbNoSpruch           As Boolean
Global gbSound              As Boolean
Global gbISDEMO             As Boolean
Global gbEKMAX              As Boolean
Global gbBONWG              As Boolean
Global gBYTEWGNR            As Byte
Global gbGiltAlsRechnung    As Boolean
Global gbEtiEan             As Boolean
Global gbEtiQuickScanM      As Boolean

Global gbKK_Visa                As Boolean
Global gbKK_EurocardMastercard  As Boolean
Global gbKK_AmericanExpress     As Boolean
Global gbKK_DinersClub          As Boolean
Global gbKK_ECKarte             As Boolean
Global gbKK_Sonstige            As Boolean

Global gbKK_AliPay              As Boolean
Global gbKK_ApplePay            As Boolean
Global gbKK_GooglePay           As Boolean
Global gbKK_PayPal              As Boolean
Global gbKK_YabandPay           As Boolean

Global gsWGart              As String
Global gsWGBEZEICH          As String
Global gbKKSCHUB            As Boolean
Global gbKOLSCHUB           As Boolean
Global gbBARZSCHUB          As Boolean
Global gbKBSCHUB            As Boolean
Global gbSPIEGEL            As Boolean
Global gbRETVK              As Boolean
Global gbZOLLmMWST          As Boolean
Global gbZOLLonlyFirstPage  As Boolean
Global gbZOLLPrintDirekt    As Boolean
Global gbFILBONI            As Boolean
Global gbBonusBNB           As Boolean
Global gbAA                 As Boolean
Global gbTagAkt             As Boolean
Global gbECTOZ              As Boolean
Global gsMeldestatus        As String
Global gbOGV                As Boolean
Global gbRGO                As Boolean
Global gbGutschnrKomplett   As Boolean
Global gbGutscheinBeiVKversteuern As Boolean
Global gbNurBonusfRunden    As Boolean
Global gbNachfragenbeiWGNohnePreis As Boolean
Global gbREGEB              As Boolean
Global gbTerminReminderSMS  As Boolean
Global gsTerminReminderstart As String
Global giGebTage            As Integer
Global gbGebAdresse         As Boolean
Global gbGesEKWert_anzeigen As Boolean
Global gbOptiStada          As Boolean
Global gbOptiStadaSpiel     As Boolean
Global gbSPY                As Boolean
Global gbAuto_Export_Artikelbestand As Boolean
Global gsServerIP           As String
Global gsServerPort         As String
Global gbSPEZRU             As Boolean
Global gbSPEZVAR            As Byte

Global gbPrintLOGO          As Boolean
Global gbBARDINA4           As Boolean
Global gbDINA4VIS           As Boolean
Global gbDINA4RECHFU        As Boolean
Global gbBARBON2            As Boolean
Global gbSparsatz           As Boolean
Global gbRabVs              As Boolean
Global gsGZBez              As String
Global gbKundRabattDeaktiv  As Boolean
Global gbJBTART             As Boolean
Global gbNOBONDRUCKER       As Boolean
Global gbKUDU               As Boolean
Global gbGEBRABK            As Boolean
Global gbbonusHerab         As Boolean

Global gbbonusausjetzt      As Boolean
Global gdbonusHerabwert     As Double
Global dBonusfaehig         As Double

Global gdbonusGutschein     As Double
Global gdBonusAusKundenhistorie As Double
Global gbBonusEinloesungHierErlaubt As Boolean



Global dwertGutverkauf      As Double
Global gsKaMail             As String
Global gsKassDatstart       As String
Global glErrtime            As Long

'IPStat
Global gbIPSTAT             As Boolean
Global gsIPMarktNr          As String

'VEDESStat
Global gbVEDESSTAT          As Boolean
Global gsVEDESMarktNr       As String

'alles für die Nacht
Global gbNacht              As Boolean
Global gbPCAus              As Boolean
Global gbEXTSICH            As Boolean
Global gsNachtstart         As String
Global giSTARTMIN           As Integer
Global giINTERV             As Integer
Global gbUPRO               As Boolean
Global gbUSTADA             As Boolean
Global gbUSTAT              As Boolean
Global gbUKDAT              As Boolean
Global gbEKDAT              As Boolean
Global gbBR                 As Boolean
Global gbKABSCH             As Boolean
Global gbUmsatzNeu          As Boolean
Global gbKSF                As Boolean
Global gbAABSCHL            As Boolean
Global gbCheckThisNetz      As Boolean

Global gbKUIBONname         As Boolean
Global gbKUIBONvorname      As Boolean
Global gbKUIBONtitel        As Boolean
Global gbKUIBONfirma        As Boolean
Global gbKUIBONplz          As Boolean
Global gbKUIBONort          As Boolean
Global gbKUIBONstrasse      As Boolean
Global gbKUIBONtel          As Boolean
Global gbKUIBONmobil        As Boolean

Global gbTAGFILT            As Boolean
Global gbARTKUM             As Boolean
Global gbARTKUM_ohneWGN     As Boolean
Global gbKK                 As Boolean
Global gbEA                 As Boolean
Global gbMitExport          As Boolean
Global gbZBONDINA4HOCH      As Boolean

Global gbLOGO1              As Boolean
Global gbLOGO2              As Boolean
Global gbLOGO3              As Boolean

Global UeberlaufFehler      As Boolean

Global FTP_Server       As String
Global FTP_User         As String
Global FTP_PassW        As String
Global giKissFtpMode    As Integer
Global gbFtpZENT        As Boolean
Global gbFilKasDat      As Boolean
Global gbFTPautomatic   As Boolean
Global gbWVNOT          As Boolean
Global gbDSL            As Boolean
Global gbPASSIVMODE     As Boolean

Global gsDatum          As String
Global gsTastText       As String
Global gsTTextB         As TextBox
Global glKalTop         As Long
Global glKalLeft        As Long

'******End Thomas ****************

Global byteSortReihen   As Byte
Global gbLokalModus     As Boolean
Global gsAnzeigeText    As String

Global gbBargeldEingabe As Boolean

Type Errormessage
    gsFehlertext     As String
    gsFormular       As String
    gsFunktion       As String
    gsDescr          As String
    gsNumber         As Long
End Type

Global Fehler As Errormessage

Type DBErrormessage
    gsBedname       As String
    gsBednr         As String
    gsDatum         As String
    gsZeit          As String
    gsPcname        As String
End Type

Global DBFehler As DBErrormessage

'*****Programmeinstellungen*******
Global glButtonHintergrund_from As Long
Global glButtonHintergrund_to As Long
Global glButtonMouseMove_Hintergrund_from As Long
Global glButtonMouseMove_Hintergrund_to As Long
Global glButtonMouseMove_Bordercolor As Long
Global glButtonBordercolor As Long
Global glButtonMouseMove_Forecolor As Long
Global glButtonForecolor As Long

Global glH1             As Long
Global glU1             As Long
Global glS1             As Long
Global gsPname          As String
Global glH2             As Long
Global glWarn           As Long
Global glLink           As Long
Global gsFont           As String
Global gsFontsize       As Integer
Global glSelBack1       As Long
Global gsUpdPfad        As String
Global gsZinPfad        As String
Global gsZoutPfad       As String
Global gsKinPfad        As String
Global gsFotoPfad       As String
Global gsWebcamPfad     As String
Global gsDabaPfad       As String
Global gsDTAPfad        As String
Global gsSpanne         As String
Global gbLekMax         As Boolean

Global gsKL_DATENBANKNAME   As String
Global gsKL_DSN             As String
Global gsKL_ADRESSE         As String
Global gsKL_BENUTZER        As String
Global gsKL_PASSWORT        As String

Global gbKL_LIVENACHRICHTEN     As Boolean
Global gbKL_LIVEBESTAND     As Boolean
Global gbKL_LIVEKVKPR       As Boolean
Global gbKL_LIVEFarbe       As Boolean
Global gbKL_LIVEGefSperr    As Boolean
Global gbKL_LIVEGUTSCHEIN   As Boolean

Global gbKL_LIVEBESTAND_DIFF     As Boolean

'Global gsMySQL_DATENBANKNAME    As String
'Global gsMySQL_ADRESSE          As String
'Global gsMySQL_BENUTZER         As String
'Global gsMySQL_PASSWORT         As String
Global gbMySQL_LIVEBESTAND      As Boolean
Global gsMySQL_PHP_SCRIPT_PFAD  As String

Global gsMySQL_BESTAND_TAB          As String
Global gsMySQL_BESTAND_INDEXSPALTE  As String
Global gsMySQL_BESTAND_SPALTE       As String

Global gbSTADAP         As Boolean
Global gbFTH            As Boolean
Global gbKVKSicher      As Boolean
Global gbNOWOCHENDATEN  As Boolean
Global gbErrPrint       As Boolean
Global gbEtiFokEan      As Boolean
Global giAufrunden      As Integer
Global giAbrunden       As Integer
Global giRundkrit       As Byte
Global gbFtpYes         As Boolean
Global gsMDEGERAET      As String
Global gsWAAGE          As String
Global gsDFU            As String
Global gbYtescanPcom    As Byte
Global giMDEPAUSE       As Integer
Global gbYteWAAGEPcom   As Byte
Global gsWeEinzMe       As String
Global gbscanmodi       As Boolean
Global gbETIKVKAE       As Boolean
Global gbWEautoGef      As Boolean
Global gbNONEGZU        As Boolean
Global gbAutoZwsp       As Boolean
Global gbETIONLYME      As Boolean
Global gbNoETIWeAusBe   As Boolean
Global gbLocalSec       As Boolean
Global gb2BONKA         As Boolean
Global gb2BONKR         As Boolean
Global gb2BONGUVK       As Boolean
Global gb2BONEA         As Boolean
Global gb2BONTermin     As Boolean

Global gb2BONUSMESS     As Boolean
Global gb2GUTVERK       As Boolean
Global gbTerminNoWarn   As Boolean
Global gb2BONKB         As Boolean
Global gb2BONVerleih    As Boolean
Global gb2BONFI         As Boolean
Global gb2BONKOLLVK     As Boolean
Global gbBONNRUNTER     As Boolean
Global gbKASSNRUNTER    As Boolean
Global gsSTERNZEICH     As String
Global glZeichenAnzahlBon As Long
Global gdBONFONTSIZE     As Double
Global gsBONFONTNAME     As String
Global gsZOLLARTBEZ     As String
Global gb2BONST         As Boolean
Global gbBonkopie       As Boolean
Global gbSaveReport     As Boolean
Global gbDSDRUCKEN      As Boolean
Global gbDSKLEIN        As Boolean
Global gbDSMeldungErfolg As Boolean
Global gbDS_GEB_DRUCKEN As Boolean
Global gdTabfak         As Double
Global gsEdeka          As String
Global gbAutoLokalModus As Boolean
Global gbAutoSYN        As Boolean
Global gbnachkomp       As Boolean
Global gbSondRab        As Boolean
Global gbKurzerStorni   As Boolean
Global gbSCHUBMB        As Boolean
Global gbGUTSCHBARAUSZAHLUNGMITUNTER As Boolean
Global gbSTORNOcheck2Bed As Boolean
Global gbOLDSTADADEL    As Boolean
Global gbyLugBe         As Byte
Global giTageZugang     As Integer
Global giTageVerkauf    As Integer
Global gdStadaPause     As Double
Global gsKassPass       As String
Global gsKUPFAD         As String
Global gsZBon           As String
Global gsZählbeleg      As String
Global gbSichernYes     As Boolean
Global gsSICHTIME       As String
Global giSICHTYP        As Integer
Global gsSicherPfad     As String
Global gsTankPfad       As String
Global gsConverterPfad  As String
Global glUPDCOUNT       As Long
Global glUPDTime        As Long
Global gbKUNDENA        As Boolean
Global gbKUNDENn        As Boolean
Global glLokalAktuZeit  As Long
Global gsWeEinzFo       As String
Global glArtNrBeg       As Long
Global gsMWST           As String
Global gbEcash          As Boolean
Global gbUnistatWeek    As Boolean
Global gbUnistatMonat   As Boolean
Global gsEPartner       As String
Global gsStatkundnr     As String
Global gbStatweekperMail As Boolean
Global gsStatZusatz     As String
Global gdateStatlast    As Date
Global gsMStatkundnr    As String
Global gdateMStatlast   As Date
Global gsiGESRAB        As Single
Global gsiKUBONUS_SCHWELLE        As Single
Global gsGESRABBEZ      As String
Global gbZugriffNew     As Boolean
Global gsProteil        As String
Global gbLibesnrSeek    As Boolean

Global gbSchreibRechnerProto As Boolean
'***Ende Programmeinstellungen****

'adt
Global gsAdtBeleg               As String
Global gsAdtVerfahren           As String
Global gADTclientId             As String
Global gADTbezahlart            As String
Global gADTioPfad               As String
Global gADTtermId               As String
Global gADTLimit                As Integer
Global AktEcashBon              As ECASHBON
Global AktBonKS                 As ECASHBON
Global AktBonB2                 As ECASHBON
Global gbADTBON                 As Boolean
Global gbADTBonDruckerStatus    As Boolean
Global gbADTVI                  As Boolean
Global gbADTAE                  As Boolean
Global gbADTDI                  As Boolean
Global gbADTEU                  As Boolean
Global gADTipAdress             As String
Global gADTport                 As String

'elPAY
Global gELPclientId             As String
Global gELPioPfad               As String
Global giELPAY_Fehler           As Integer

'ZVT
Global gZVTclientId             As String
Global gZVTioPfad               As String
Global gZVTPName                As String
Global gZVTPTitel               As String
Global giZVT_Fehler             As Integer
Global gZVTDruckVar             As String
Global gZVTTimeOut              As Long

'ZVT2
Global gZVT2_IP                 As String
Global gZVT2_Lizenz             As String
Global gZVT2_Port               As String
Global giZVT2_Fehler            As Integer
Global gbZVT2_KBDrucken         As Boolean
Global gbZVT2_Kartenwahl        As Boolean

Global gbZVT2_HBDrucken         As Boolean
Global giZVT2_TIMEOUT           As Integer
Global gsZVT2_VirtuellID        As String
Global gbZVT2_vTIDdrucken       As Boolean


Global bErmachtigung As Boolean
Global bUnterschrift As Boolean

Global gsAnzeige00a     As String
Global gsUpdDatName     As String
Global gbSichernHeut    As Boolean
Global gbGraph As Boolean
Global giAnzFil  As Integer
Global giFilNrS()  As Integer
Global gbGutschUNDlastschrift As Boolean
Global gdGutLastRest As Double
Global gcDivKosmetik As String
Global gbDivKosmetik As Boolean
Global gdWertDEM As Double
Global gcRueckgeld As String
Global gsETILS  As String
Global gbEtiExArtikel As Boolean
Global glEtiExArtikel_linr As Long

Global gbPreisAender As Boolean
Global gRegister As REGISTERINFO_
Global gDtaASatz As DTA_ASATZ_
Global gDtaCSatz As DTA_CSATZ_
Global gDtaESatz As DTA_ESATZ_
Global gDtaBegleit As DTA_BEGLEIT_
Global gVerbindung As VERBINDUNG_

Global gF2Prompt As F2PROMPT_

Global gLayout() As LAYOUT_

Global glWGTaste(0 To 19) As Long
Global gWarenGruppe(18 To 171) As WARENGRU_
Global DlgZugriff(0 To 49) As DLG_ZUGRIFF
Global gFirma As FIRMA_
Global gZeiten(1 To 21) As ZEITEN_
Global gECKarte As ECKARTE_
Global gKasseIni As KASSEINI_

Global gbDisplay                As Boolean
Global gbGutsch                 As Boolean
Global gbGutschein              As Boolean
Global gbDisplaySeriell         As Boolean
Global gbWarenGruppe            As Boolean
Global gbDebug                  As Boolean
Global gbTagesAbschluss         As Boolean
Global gbZwangsKdNr             As Boolean
Global giAnzBonus_Erreicht      As Integer
Global glBONUSAA_Artnr          As Long
Global glBONUSNRE_Artnr         As Long
Global gbABOPLUS                As Boolean
Global gbNumTaste               As Boolean
Global gbLadeCom                As Boolean
Global gbDrucken                As Boolean
Global gbStornoErlaubt          As Boolean
Global gbRabatt As Boolean
Global gbLeiste2Start As Boolean
Global gbIdentUser As Boolean
Global gbGrossLief As Boolean
Global gbDBMod As Boolean
Global gbAbschlussNummer As Boolean
Global gbAbschlussDatum As Boolean
Global gbAGNAusw As Boolean
Global gbARTKUMWGN As Boolean
Global gbKUMSUM As Boolean
Global gbSTAMDA As Boolean
Global gbMB As Boolean
Global gbabbruchG As Boolean

Global gbKUWAHLfbimmer As Boolean
Global gbKUWAHLROT As Boolean
Global gbKUWAHLGESPERRTROT As Boolean
Global gbKUWAHLMAIL As Boolean
Global gbKUBONUS As Boolean
Global gbNoKUBONUS_wenn_Art_and_Ges_rab As Boolean
Global gbMitPreis As Boolean
Global gbAUSBLDU As Boolean
Global gbAUSBLSH As Boolean
Global gbAUSBLLS As Boolean
Global gbNoGrafik As Boolean
Global gbOpenSchubRetoure As Boolean
Global giBARGELDART As Integer
Global gbMBBLOCKFrage As Boolean
Global gsSperrFrage As String
Global gbNoBonGu As Boolean
Global gbBonGu2J As Boolean
Global gbNoBonPÄ As Boolean
Global gdRESTGU As Double
Global giFILALI As Integer
Global giFarbebeiPark As Integer
Global glRRArtnr As Long
Global glBonusGrenzeArtnr As Long
Global glBonusAuszahlungArtnr As Long
Global glECAuszahlArtnr As Long
Global gdSCHWELLEWK As Double
Global glAutoKundnrforKundBest As Long
Global glAutoAusSchFiliale As Long
Global glTageVorTermin As Long
Global gsAbrunden As String
Global gsFARBKASSE As String
Global gsECBILD As String
Global glGSArtnr As Long
Global glBaganzArtnr As Long
Global glBaganzAR As Long

Global glZehnProzLinr As Long
Global glZehnProzArtnr As Long

Global glLieblingArtnr As Long
Global glLieblingAR As Long

Global gdWarenkorbWert As Double
Global gdWarenkorbGR As Double

Global glPrimLinr As Long
Global glZeitungsLinr As Long
Global glPaketLinr As Long
Global glSpezFotoartikel As Long
Global glSpezLottoauszahlartikel As Long
Global gsSpezArtikel As String
Global gsRabattAusnahmeArtikel As String
Global gskPW As String
Global gsSpezBontext As String
Global gsSpezBontext2 As String
Global gsSpezBontext3 As String
Global gsSpezBontextU As String
Global gsSpezBonArtRab As String
Global gdZeitungsSpanne As Double

Global gbKUBONUS_WENN As Boolean

Global giDisplaySeriellComPort As Integer
Global giPreisKz As Integer
Global giAFCBUCH As Integer

Global gcMwSt As String
Global gcUserName As String
Global gcAnwender As String
Global gcPass As String
Global glLevel As Long

Global gcIdentUserName As String
Global gcIdentPass As String
Global gcIdentBedienerNr As String
Global glIdentLevel As Long

Global gcIdentStornoBedienerNr1 As String
Global gcIdentStornoBedienerNr2 As String

Global gcDlgTitel As String
Global gcNetzLW As String

Global gcPfad As String
Global gcDBPfad As String
Global gcNetPfad As String
Global gcSysPfad As String
Global gcBediener As String
Global gcBedienerNr As String
Global gcALTBedienerNr As String
Global gcKreditKarte As String
Global gcKreditKarte2 As String
Global gcPasswort As String
Global gckundnr As String

Global gckuVorname As String
Global gckuname As String

Global gcKundNrDaten As String
Global gcLadeCom As String
Global gcSchneiden As String
Global gcLade As String
Global gcBild As String
Global gcBarCode As String
Global gcInit As String
Global gcKuerzel As String
Global gcArtNrFiliale As String
Global gcDisplay As String
Global gcReNr As String
Global gcRePreisKz As String
Global gcRueckGutsch As String
Global gcKassenDatei As String

Global gsTerminalid As String

Global gdBase As Database
Global gdApp As Database
Global gdbNet As Database
Global gdbFiliale As Database
Global dbPrintAusKissdata As Database

Global gcSMTP_SERVER As String
Global gcSMTP_USER As String
Global gcSMTP_PW As String
Global gcSMTP_PORT As String
Global gbSMTP_SSL As Boolean

Global gcKasNum As String
Global gcFilNr As String
Global gbFilNr As Boolean
Global gbLizenz As Boolean
Global gbLizenzINDI As Boolean
Global glSelect As Long
Global glPosStorno As Long
Global glGarantienummer As Long
Global glGutschNr As Long
Global giDlgZustand As Integer
Global giAndersZahlung As Integer
Global glLiNr As Long
Global giWochendat As Integer
Global giDialog As Integer
Global giSortierung As Integer
Global giErsetzen As Integer
Global giKassenDatei As Integer

Global gbNeuerSatz As Boolean
Global gbAutomaticClick As Boolean
Global gbFertig As Boolean
Global gbAPI As Boolean
Global gbRegister As Boolean
Global gbNetzLW As Boolean
Global gbBonusKunde As Boolean
Global gdBonNr As Double
Global gdBonusGrenze As Double
Global gdBonusGutscheinBeiGrenze As Double

Global gdSummeLAST As Double
Global gdSummeKK As Double
Global gdSumme As Double
Global gdGegeben As Double

Global gdGegebenFW As Double
Global gsFWK As String

Global gdZurueck As Double
Global giZahlArt As Integer
Global gcZahlMittel As String

Global gcTag As String
Global gcWochentag(0 To 7) As String
Global gcMonat(1 To 12) As String
Global gcBonText(0 To 12) As String
Global gcZahlArt(0 To 53)   As String
Global gsLinr               As String
Global gsArtikelArray(0 To 9) As String

Global gsVMPadresse As String
Global gsVMPbetreff As String
Global gsVMPKdNr As String
Global gsVMPzLinr As String
Global gsVMPEndung As String
Global gsVMPArt As String

Global gbSHOPARTIKEL As Boolean
Global gbPlusEAN As Boolean
Global gbPlusBezeich As Boolean
Global gbPlusShopPreis As Boolean
Global gbMITUEBERSCHRIFT As Boolean
Global gbNMB    As Boolean
Global gbEXNOR    As Boolean
Global gbBL    As Boolean
Global gsBLKENNUNG   As String
Global gsDATEIENDUNG As String
Global gsFELDTRENNER As String

Global gclinr11(1 To 999) As String
Global gclinr As String
Global gcKunden() As String
Global glAnzKunden As Long
Global glStatusKunden As Long
Global gdRechner(0 To 5) As Double
Global gbTransfer(1 To 15) As Boolean
Global glAnzLiNr As Long
Global gbBonDruck As Boolean
Global gbLastSchriftEnde As Boolean

Global gcBedNr As String
Global gcBonDrucker As String
Global gcListenDrucker As String
Global gcEtikettenDrucker As String
Global gcFaxDrucker As String
Global gcDefaultDrucker As String
Global gcGutscheinDrucker As String

Global glBestandAlt As Long
Global glBestandNeu As Long


Global gcKundenNr As String

Global gcZeitBlock As String
Global gcStartZeit As String
Global gcEndeZeit As String

Global giKummListe As Integer

Global gcDLL1 As String
Global gcDLL2 As String

Global lDBVersion As Long
Global WKVersion As Long

Global gcWaehrung As String

Global gdMWStV As Double
Global gdMWStE As Double
Global gdMWStO As Double

Global gbDosKasse As Boolean

'Konstanten

Global glfarbe(0 To 9) As String
Global glfarbe2(0 To 9) As String

Global Const gdDEMFaktor As Double = 1.95583
Global Const gdATSFaktor As Double = 13.7603
Global Const gdNLGFaktor As Double = 2.1735

Global Const gdCHFFaktor As Double = 1
Global Const gsPasswort   As String = "Kiss2005"
Global Const gsGDPdU_Passwort   As String = "§14UStG"
Global Const gsKASSBON_Passwort As String = "MXGZ"
Global Const gcMASTER As String = "KISS2000"
Global Const gcMASTERUSER As String = "KISSLITE"


Global Const gcRegDatei As String = "CMCTLS32.DLL"
Global Const gcStornoPW As String = "X2§$)4QX2FLPSSE31"

Global Const gcNUM As String = "1234567890"
Global Const gcUPPER As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Global Const gcLower As String = "abcdefghijklmnopqrstuvwxyz"


Global Const giUPD As Integer = 0
Global Const giNEU As Integer = 1

Global Const giKREDIT As Integer = 5
Global Const giKOLLEGE As Integer = 47



Global Const glEUROTAG As Long = 37257
'API-Konstanten

Global Const WM_SETTINGCHANGE = &H1A
Global Const HWND_BROADCAST = &HFFFF&
Global Const WM_WININICHANGE = &H1A

Global Const WM_USER = &H400
Global Const LB_SETHORIZONTALEXTENT = WM_USER + 21

'internet Email
Global Const DIAL_UNATTENDED = &H8000
Global Const DIAL_FORCE_ONLINE = 1
Global Const DIAL_FORCE_UNATTENDED = 2

Global Const RAS95_MaxEntryName = 256

Type RASENTRYNAME95
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type

Global ConID        As Long
Global ConName      As String

'Autocomplete
Global Const WM_SETREDRAW = &HB

Global Const CB_OKAY = 0
Global Const CB_ERR = -1
Global Const CB_ERRSPACE = -2

Global Const CB_FINDSTRING = &H14C
Global Const CB_SELECTSTRING = &H14D
Global Const CB_FINDSTRINGEXACT = &H158

Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As RASStatusType) As Long
Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As RASType, lpcb As Long, lpcConnections As Long) As Long
Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
        
Global Const RAS_MaxEntryName = 256
Global Const RAS_MaxDeviceType = 16
Global Const RAS_MaxDeviceName = 32 'oder 128
Global Const Max_Fill = 96

'Global Const RAS95_MaxDeviceName = 128

Type RASStatusType
  dwSize As Long
  RasConnState As Long
  dwError As Long
  szDeviceType(RAS_MaxDeviceType) As Byte
  szDeviceName(RAS_MaxDeviceName) As Byte
End Type

Public Type RASType
  dwSize As Long
  hRasCon As Long
  szEntryName(RAS_MaxEntryName) As Byte
  szDeviceType(RAS_MaxDeviceType) As Byte
  szDeviceName(RAS_MaxDeviceName) As Byte
End Type

'Type RASType
'    dwSize As Long
'    hRasConn As Long
'    szEntryName(RAS_MaxEntryName) As Byte
'    szDeviceType(RAS_MaxDeviceType) As Byte
'    szDeviceName(RAS_MaxDeviceName) As Byte
'    dwFill(Max_Fill) As Byte
'End Type

Type Systeminfo
    dwOEMID As Long
    dwPagesize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Global Const VER_PLATFORM_WIN32s = 0
Global Const VER_PLATFORM_WIN32_WINDOWS = 1
Global Const VER_PLATFORM_WIN32_NT = 2

'sound
Declare Function PlaySound Lib "winmm.dll" Alias _
        "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal _
        uFlags As Long) As Long

Dim rs$, cb&
'Ende

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function RegCreateKey& Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey&, ByVal lpszSubKey$, lphKey&)
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS As Long = &H80000003

Global Const REG_SZ = 1
Global Const REG_BINARY = 3
Global Const REG_DWORD As Long = 4&

Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, ByVal lpszConiID As String, ByVal dwFlags As Long, ByRef hCon As Long, ByVal dwReserved As Long) As Long
Declare Function InternetHangUp Lib "wininet.dll" (ByVal hCon As Long, ByVal dwReserved As Long) As Long

Declare Function RasEnumEntries Lib "rasapi32.dll" _
        Alias "RasEnumEntriesA" (ByVal Reserved$, ByVal _
        lpszPhonebook$, lprasentryname As Any, lpcb As Long, _
        lpcEntries As Long) As Long
        
'internet Email ende
'Alle Task
Declare Function GetDesktopWindow Lib "user32" () _
        As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd _
        As Long, ByVal wCmd As Long) As Long
        
Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hwnd As Long, ByVal wIndx As _
        Long) As Long
        
Declare Function GetWindowTextLength Lib "user32" _
        Alias "GetWindowTextLengthA" (ByVal hwnd As Long) _
        As Long
        
Declare Function GetWindowText Lib "user32" Alias _
        "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString _
        As String, ByVal cch As Long) As Long
               
Declare Function GetParent Lib "user32" (ByVal hwnd _
        As Long) As Long
        
Declare Function GetWindowThreadProcessId Lib "user32" _
        (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Declare Function ShowWindow Lib "user32" ( _
  ByVal hwnd As Long, _
  ByVal nCmdShow As Long) As Long


Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5
Global Const GW_MAX = 5
Global Const WM_CLOSE = &H10
Global Const SW_MINIMIZE = 6 ' Minmiert das Fenster
Global Const SW_HIDE = 0 '' Versteckt das Fenster
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOWNORMAL = 1
Global Const GWL_STYLE = (-16)

Global Const WS_VISIBLE = &H10000000
Global Const WS_BORDER = &H800000
Global ITask As Integer

'Alle Task Ende

Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) _
        As Long

Declare Function CloseHandle Lib "kernel32" (ByVal _
        hObject As Long) As Long
   
Declare Function OpenProcess Lib "kernel32" (ByVal _
        dwDesiredAccess As Long, ByVal bInheritHandle As _
        Long, ByVal dwProcessID As Long) As Long
        
Global Const INFINITE = -1&
Global Const SYNCHRONIZE = &H100000
        


Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" _
    (lpNetResource As NETRESOURCE, ByVal lpPassword As String, _
    ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
    "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, _
    ByVal fForce As Long) As Long
    

  ' Für WNetAddConnection2 und WNetCancelConnection2 (teilweise)
  Global Const CONNECT_UPDATE_PROFILE = &H1 ' Beim nächsten Login des Benutzers wieder verbinden
  Global Const RESOURCETYPE_DISK = &H1 ' Einzubindende Netzressource ist ein Laufwerk
  Global Const NO_ERROR = 0 ' Erfolgreiche Abarbeitung
  Global Const ERROR_ACCESS_DENIED = 5& ' Zugriff verweigert
  Global Const ERROR_ALREADY_ASSIGNED = 85& ' Lokaler Laufwerksbuchstabe ist bereits verbunden
  Global Const ERROR_BAD_DEV_TYPE = 66& ' Lokales Laufwerk hat anderen Typ als die Netzwerk-Ressource
  Global Const ERROR_BAD_DEVICE = 1200& ' Die Laufwerksbezeichnung in .lpLocalName ist ungültig
  Global Const ERROR_BAD_NET_NAME = 67& ' Die Ressourcenbezeichnung in .lpRemoteName ist ungültig
  Global Const ERROR_BAD_PROFILE = 1206& ' Benutzerprofil ist ungültig
  Global Const ERROR_BAD_PROVIDER = 1204& ' Angegebener Provider (.lpProvider) ist ungültig
  Global Const ERROR_BUSY = 170& ' Netzwerk ist beschäftigt - später erneut versuchen
  Global Const ERROR_CANNOT_OPEN_PROFILE = 1205& ' Persistente Verbindung nicht möglich, _
                da das Benutzer-Profil nicht geöffnet werden konnte
  Global Const ERROR_INVALID_PASSWORD = 86& ' Username/Password-Kombination ist falsch
  Global Const ERROR_NO_NET_OR_BAD_PATH = 1203& ' Keine Verbindung zum Netzwerk oder Pfad nicht erkannt
  Global Const ERROR_NO_NETWORK = 1222& ' Es ist keine Netzwerkunterstützung installiert
  ' zusätzlich für WNetCancelConnection2
  Global Const ERROR_DEVICE_IN_USE = 2404& ' Die Ressource kann nicht getrennt werden
  Global Const ERROR_NOT_CONNECTED = 2250& ' Es bsteht keine Verbindung zur angegebenen Ressource
  Global Const ERROR_OPEN_FILES = 2401& ' Das Laufwerk ist in Benutzung und bForce ist False
  
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'Deklaration: Globale Form API-Funktionen
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

  

  
'API-Deklarationen
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As Systeminfo)
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetVersionEx Lib "kernel32" Alias _
        "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) _
        As Long

Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
 
Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

'für Win.ini
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" _
        (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long) As Long

Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" _
        (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
'für Win.ini Ende

'für x.ini
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal _
        lpFileName As String) As Long
        
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
        
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
        (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName _
        As String) As Long
        
'für x.ini Ende

'Satzbeschreibung für Listbox KASSE

'001 - 005  MENGE
'006 - 006  LEER oder STERN (*) für Rabattfähig
'007 - 012  ARTIKELNUMMER
'013 - 013  LEER
'014 - 048  ARTIKELBEZEICHNUNG
'049 - 049  LEER
'050 - 058  EINZELPREIS (nach Abzug Rabatt)
'059 - 059  LEER
'060 - 068  GESAMTPREIS (Einzelpreis x Menge)
'069 - 071  LEER
'072 - 072  MWSTKZ
'073 - 073  LEER
'074 - 082  LISTENPREIS bzw. SONDERPREIS
'083 - 083  LEER
'084 - 092  ERMAESSIGUNGSBETRAG BEI ARTERM
'093 - 093  LEER
'094 - 102  ERZIELTER VERKAUFSPREIS
'103 - 103  LEER
'104 - 112  MWSTBETRAG VOLLE MWST
'113 - 113  LEER
'114 - 122  MWSTBETRAG ERM MWST
'123 - 123  LEER
'124 - 126  ARTIKELRABATT IN %
'127 - 127  LEER
'128 - 136  VKPREIS
'137 - 137  LEER
'138 - 146  RESTBESTAND
'147 - 147  LEER
'148 - 150  BEDIENERNUMMER
'151 - 151  LEER
'152 - 152  RABATTFAEHIG (0 = Ja, 1 = Nein)
'153 - 153  LEER
'154 - 154  BONUS_OK (J/N)
'155 - 155  LEER
'156 - 156  LEER
'157 - 157  UMS_OK (J/N)
'158 - 175  ZEITSCHRIFTEN-ID (EXTEND)
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Declare Function GetLocaleInfo Lib "kernel32" _
        Alias "GetLocaleInfoA" (ByVal Locale As Long, _
        ByVal LCType As Long, ByVal lpLCData As String, _
        ByVal cchData As Long) As Long
        
Global Const LOCALE_SSHORTDATE = &H1F
Global Const LOCALE_SLONGDATE = &H20

Declare Sub keybd_event Lib "user32" (ByVal bVk As _
        Byte, ByVal bScan As Byte, ByVal dwFlags As Long, _
        ByVal dwExtraInfo As Long)

Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_LWIN = &H5B


Global WshShell As Object
  
