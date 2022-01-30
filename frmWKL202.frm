VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL202 
   Caption         =   "Schapfl Import"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL202.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picprogress 
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   9315
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   7320
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   6
      Left            =   9600
      TabIndex        =   3
      Top             =   7080
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "importieren"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorDisabled=   15398133
      BackColorFrom   =   16514300
      BackColorTo     =   15462640
      BackColorCheckedFrom=   15462640
      BackColorCheckedTo=   16514300
      BackColorDownFrom=   12700881
      BackColorDownTo =   15659506
      BackColorHoverFrom=   16514300
      BackColorHoverTo=   15462640
      BorderColor     =   7617536
      BorderColorDisabled=   12240841
      BorderColorFocus=   14986635
      BorderColorHover=   3913721
      ForeColorDisabled=   9609633
      MenuBackColor   =   16448250
      MenuBackColorChecked=   7323903
      MenuBackColorHover=   10935807
      MenuBorderColor =   8388608
      MenuCheckMarkColorFrom=   16514300
      MenuCheckMarkColorTo=   15462640
      MenuForeColor   =   -2147483640
      MenuForeColorHover=   -2147483640
      ButtonStyle     =   2
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   6720
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "- Access Datenbank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   10095
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Was benötigen Sie für den Datenimport?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11535
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7800
      Width           =   9255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Schapfl Import"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmWKL202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad As String
    Dim sdbPfad As String

    Select Case Index
        Case 0
            Unload frmWKL202
        Case 6
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Wo ist die Schapfl - Datei?"

                .Filter = "Access - Dateien (*.mdb)|zentrale.mdb"
                .ShowSave

                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
            End With
            
            SchapflImport Label1(4), sPfad
    End Select

err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Schapfl Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SchapflImport(lblx As Label, cPfad As String)
On Error GoTo LOKAL_ERROR

    lblx.Caption = TimeValue(Now) & ": Tabellen werden in den Speicher geladen...": lblx.Refresh
    
    Dim db          As Database
    Dim cPfad1      As String
    Dim sSQL        As String
    Dim rsrsZ       As DAO.Recordset
    Dim rsRSQ       As DAO.Recordset
    Dim rsrs        As DAO.Recordset
    Dim ineueBednr  As Integer
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    
    Set db = OpenDatabase(cPfad & "\zentrale.mdb", False, False)
    
    
    
'    'Gutscheine keine enthalten
'
'    loeschNEW "Gutscheine", gdBase
'    TransferTab db, cPfad1 & "kissdata.mdb", "Gutscheine"
'
'
'
    sSQL = "Delete * from Gutsch "
    gdBase.Execute sSQL, dbFailOnError
'
'
'    sSQL = "Insert into Gutsch Select "
'    sSQL = sSQL & " gutnr as gutschnr   "
'    sSQL = sSQL & ", ausstell as DAT_AUSG "
'    sSQL = sSQL & ", einloes as DAT_EINL "
'    sSQL = sSQL & ", kdnr as Kundnr "
'    sSQL = sSQL & ", ausbetrag as wert "
'    sSQL = sSQL & ", 1 as FILIALE "
'    sSQL = sSQL & " from tblGutschein  "
'    gdBase.Execute sSQL, dbFailOnError
    
    'Firma
    loeschNEW "orgaeinheit", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "orgaeinheit"
    
   
    
    sSQL = "Delete * from Firma "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Firma Select "
    sSQL = sSQL & " Name,Strasse,Plz,Ort "
    sSQL = sSQL & " from orgaeinheit  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    'BONTEXT
    loeschNEW "zeile", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "zeile"
    
    sSQL = "Delete * from BONTEXT "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 0 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'J' "
    sSQL = sSQL & " and Position = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 1 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'J' "
    sSQL = sSQL & " and Position = 1 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 4 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'J' "
    sSQL = sSQL & " and Position = 2 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 12 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'J' "
    sSQL = sSQL & " and Position = 3 "
    gdBase.Execute sSQL, dbFailOnError
        
    'Fußzeile
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 2 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'N' "
    sSQL = sSQL & " and Position = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 3 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'N' "
    sSQL = sSQL & " and Position = 1 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 5 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'N' "
    sSQL = sSQL & " and Position = 2 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BONTEXT Select "
    sSQL = sSQL & " 6 as ZEILENNR "
    sSQL = sSQL & " ,Text as ZEILENTEXT "
    sSQL = sSQL & " from zeile where kopfzeile = 'N' "
    sSQL = sSQL & " and Position = 3 "
    gdBase.Execute sSQL, dbFailOnError
        
    
    'Ende Bontext
    
    
    'Filialen
    sSQL = "Delete * from Filialen "
    gdBase.Execute sSQL, dbFailOnError
    
    
'    sSQL = "Insert into Filialen Select "
'    sSQL = sSQL & " Name1 as FilialName   "
'    sSQL = sSQL & ", 1 as FILIALNR "
'    sSQL = sSQL & " from tblParameter  "
'    gdBase.Execute sSQL, dbFailOnError
    
    
    
    

    
    
    
    'Lieferanten
    
    loeschNEW "Lieferant", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "Lieferant"
    
    sSQL = "Delete from LISRT"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LISRT Select "
    sSQL = sSQL & " LiefNrInt as LINR  "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", plz "
    sSQL = sSQL & ", name as liefbez "
    sSQL = sSQL & ", Ucase(left(name,5)) as  KUERZEL "
    sSQL = sSQL & ", Ort as STADT "
    sSQL = sSQL & ", Telefon as TEL "
    sSQL = sSQL & ", FAX "
    sSQL = sSQL & ", kdnr as KUNDNR "
    sSQL = sSQL & ", EMAIL "

    sSQL = sSQL & " from Lieferant  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    'Bediener
    
    sSQL = "Delete * from Bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Bedname (BEDNU,BEDNAME,BEDIENER) values "
    sSQL = sSQL & " (1,'Frau Brnas', 9) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
      
    'Kunden
    
'    sSQL = "Delete * from Kunden"
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Insert into Kunden Select "
'    sSQL = sSQL & " Kdnrorg as KUNDNR  "
'    sSQL = sSQL & ", anrede "
'    sSQL = sSQL & ", Ucase(left(anrede, 5)) as kuerzel "
'    sSQL = sSQL & ", strasse "
'    sSQL = sSQL & ", plzstr as plz "
'    sSQL = sSQL & ", name1 as name "
'    sSQL = sSQL & ", name2 as notizen "
'    sSQL = sSQL & ", Ort as STADT "
'    sSQL = sSQL & ", Telefon as TEL "
'    sSQL = sSQL & ", Telefax as FAXNR "
'    sSQL = sSQL & ", rgrab as Rabatt "
'    sSQL = sSQL & ", EMAIL "
'    sSQL = sSQL & " from tblAdressen  "
'    gdBase.Execute sSQL, dbFailOnError
    
    'Artikelgruppe
    
    loeschNEW "warengruppe", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "warengruppe"
    
    sSQL = "Delete * from AGNDBF"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "select * from AGNDBF "
    Set rsrsZ = gdBase.OpenRecordset(sSQL)
    
    
    sSQL = "select * from warengruppe  "
    Set rsRSQ = gdBase.OpenRecordset(sSQL)

    If Not rsRSQ.EOF Then
        rsRSQ.MoveFirst
        Do While Not rsRSQ.EOF
            rsrsZ.AddNew
            
            
            If Not IsNull(rsRSQ!WGR) Then
                If IsNumeric(rsRSQ!WGR) Then
                    rsrsZ!AGN = rsRSQ!WGR
                    
                    If Not IsNull(rsRSQ!BEZEICHNUNG) Then
                        rsrsZ!AGTEXT = Left(rsRSQ!BEZEICHNUNG, 30)
                    Else
                        rsrsZ!AGTEXT = ""
                    End If
                    
                End If
            End If
           
            rsrsZ.Update
            rsRSQ.MoveNext
        Loop
    End If

    rsRSQ.Close
    rsrsZ.Close
    
    sSQL = "Delete * from AGNDBF where agn = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from AGNDBF where agn is null "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "PGNDBF", gdBase

    sSQL = "Create Table PGNDBF ("
    sSQL = sSQL & " PGN integer "
    sSQL = sSQL & " , PGNBEZEICH Text(35) "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Umsatz wird geladen...": lblx.Refresh
    
    loeschNEW "Abschluss", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "Abschluss"
    
    SpalteAnfuegenNEW "Abschluss", "ADATE", "Datetime", gdBase
    
    sSQL = "Update Abschluss set ADATE = left(Zeitpunkt,10)"
    gdBase.Execute sSQL, dbFailOnError
    
    'Umsatz
    sSQL = "Delete * from Umsatz"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Umsatz Select "
    sSQL = sSQL & " adate as Datum   "
    sSQL = sSQL & " , sum(bruttoges) as umsg1 "
    sSQL = sSQL & " , sum(umskz0) as umso1 "
    sSQL = sSQL & " , sum(umskz9) as umse1 "
    sSQL = sSQL & " , sum(umskz4) as umsv1 "
    sSQL = sSQL & " from Abschluss group by adate "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    
    sSQL = "Update Umsatz Set kunz1 = 0,kred1 = 0, ekpr1 = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set umso1 = 0 where umso1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set umse1 = 0 where umse1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set umsv1 = 0 where umsv1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set umsg1 = 0 where umsg1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artlief wird geladen...": lblx.Refresh
    
    
    
    'Artlief
    
    loeschNEW "lieferantenartikel", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "lieferantenartikel"
    
   
    
    sSQL = "Delete * from Artlief"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    sSQL = "Insert into Artlief Select "
    sSQL = sSQL & " artid as Artnr   "
    sSQL = sSQL & ", LIEFNRINT as LINR "
    sSQL = sSQL & ", EKPREIS as LEKPR "
    sSQL = sSQL & ", 1 as MINMEN "
    sSQL = sSQL & ", LIEFARTNR as LIBESNR "
    sSQL = sSQL & ", 'N' as RKZ "
    sSQL = sSQL & " from lieferantenartikel  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artlief Set LINR = 500000 where LINR is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artlief Set LEKPR = 0 where LEKPR is null"
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Delete * from Artlief where Artnr is null"
'    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel wird geladen...": lblx.Refresh
    
    
    
    'Artikel
    
    loeschNEW "tblArtikel", gdBase
'    TransferTab db, cPfad1 & "kissdata.mdb", "tblProdLieferstamm"
    
    Dim sZiel As String
    sZiel = cPfad1 & "kissdata.mdb"
    
    sSQL = "Select Artikel.* into tblArtikel IN '" & sZiel & "' from Artikel "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from Artikel"
    gdBase.Execute sSQL, dbFailOnError
    
    

    
    sSQL = "Insert into Artikel Select "
    sSQL = sSQL & " artid as Artnr "
    sSQL = sSQL & " ,wgr as agn "
    sSQL = sSQL & " from tblArtikel "
    gdBase.Execute sSQL, dbFailOnError
    
'    loeschNEW "F_Artikel", gdBase
'    sSQL = "Select artnr into F_Artikel from Artlief where not artnr in (Select artnr from artlief)"
'    gdBase.Execute sSQL, dbFailOnError
    
    
    'MWST aus Warengruppe
    sSQL = "Update Artikel a inner join Warengruppe w on a.agn = w.wgr "
    sSQL = sSQL & " set a.MWST = 'V'   "
    sSQL = sSQL & " where w.MWSTID = 4   "
    gdBase.Execute sSQL, dbFailOnError
    
    'MWST aus Warengruppe
    sSQL = "Update Artikel a inner join Warengruppe w on a.agn = w.wgr "
    sSQL = sSQL & " set a.MWST = 'E'   "
    sSQL = sSQL & " where w.MWSTID = 9   "
    gdBase.Execute sSQL, dbFailOnError
    
    'MWST aus Warengruppe
    sSQL = "Update Artikel a inner join Warengruppe w on a.agn = w.wgr "
    sSQL = sSQL & " set a.MWST = 'O'   "
    sSQL = sSQL & " where w.MWSTID = 0   "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'noch Bezugseinheit in MINMEN
    sSQL = "Update Artlief a inner join tblArtikel t on a.artnr = t.artid "
    sSQL = sSQL & " set a.minmen = t.bezugseinheit   "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Strichcode wird geladen...": lblx.Refresh
    
    
    loeschNEW "strichcode", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "strichcode"
    
    sSQL = "Create index eanid on strichcode(eanid) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    loeschNEW "strichcodeTop1", gdBase
    
    sSQL = "Select min(eanid) as eanNr"
    sSQL = sSQL & " , artid as Artnr   "
    sSQL = sSQL & " into strichcodeTop1 from strichcode  "
    sSQL = sSQL & " group by artid "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create index eannr on strichcodeTop1(eannr) "
    gdBase.Execute sSQL, dbFailOnError
    
    SpalteAnfuegenNEW "strichcodeTop1", "EAN", "Text(13)", gdBase
    
    SpalteAnfuegenNEW "strichcodeTop1", "KVKPR1", "Double", gdBase
    SpalteAnfuegenNEW "strichcodeTop1", "BEZEICH", "Text(35)", gdBase
    SpalteAnfuegenNEW "strichcodeTop1", "INHALT", "Double", gdBase
    SpalteAnfuegenNEW "strichcodeTop1", "INHALTBEZ", "Text(3)", gdBase
    
    
    sSQL = "Update StrichcodeTop1 a inner join Strichcode s on  a.eanNr = s.eanid "
    sSQL = sSQL & " set a.ean = s.strichcode "
    sSQL = sSQL & " , a.KVKPR1 = s.VKPREIS "
    sSQL = sSQL & " , a.BEZEICH = s.BEZEICHNUNG "
    
    sSQL = sSQL & " , a.INHALT = s.MENGENEINHEIT "
    sSQL = sSQL & " , a.INHALTBEZ = Ucase(Left(s.MENGENTYP,3)) "
    
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Artikel a inner join StrichcodeTop1 s on a.artnr = s.artnr "
    sSQL = sSQL & " set a.ean = s.ean "
    sSQL = sSQL & " , a.KVKPR1 = s.KVKPR1 "
    sSQL = sSQL & " , a.BEZEICH = s.BEZEICH "
    
    sSQL = sSQL & " , a.INHALT = s.INHALT "
    sSQL = sSQL & " , a.INHALTBEZ = s.INHALTBEZ "
    
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete * from strichcode where eanid in (Select eannr from strichcodeTop1) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'und nochmal für EAN2
    loeschNEW "strichcodeTop1", gdBase
    
    sSQL = "Select min(eanid) as eanNr"
    sSQL = sSQL & " , artid as Artnr   "
    
    sSQL = sSQL & " into strichcodeTop1 from strichcode  "
    sSQL = sSQL & " group by artid "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create index eannr on strichcodeTop1(eannr) "
    gdBase.Execute sSQL, dbFailOnError
    
    SpalteAnfuegenNEW "strichcodeTop1", "EAN", "Text(13)", gdBase
    
    sSQL = "Update StrichcodeTop1 a inner join Strichcode s on  a.eanNr = s.eanid "
    sSQL = sSQL & " set a.ean = s.strichcode "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Artikel a inner join StrichcodeTop1 s on a.artnr = s.artnr "
    sSQL = sSQL & " set a.ean2 = s.ean "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete * from strichcode where eanid in (Select eannr from strichcodeTop1) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'und nochmal für EAN3
    loeschNEW "strichcodeTop1", gdBase
    
    sSQL = "Select min(eanid) as eanNr"
    sSQL = sSQL & " , artid as Artnr   "
    
    sSQL = sSQL & " into strichcodeTop1 from strichcode  "
    sSQL = sSQL & " group by artid "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create index eannr on strichcodeTop1(eannr) "
    gdBase.Execute sSQL, dbFailOnError
    
    SpalteAnfuegenNEW "strichcodeTop1", "EAN", "Text(13)", gdBase
    
    sSQL = "Update StrichcodeTop1 a inner join Strichcode s on  a.eanNr = s.eanid "
    sSQL = sSQL & " set a.ean = s.strichcode "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Artikel a inner join StrichcodeTop1 s on a.artnr = s.artnr "
    sSQL = sSQL & " set a.ean3 = s.ean "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete * from strichcode where eanid in (Select eannr from strichcodeTop1) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    'und für den Rest in die ARTEAN_K (Artnr,EAN)
    
    sSQL = "Delete * from ARTEAN_K "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ARTEAN_K select "
    sSQL = sSQL & "  artid as Artnr   "
    sSQL = sSQL & "  ,strichcode as ean   "
    sSQL = sSQL & "  from strichcode  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'restliche Standards
    sSQL = "Update Artikel a inner join artlief f on a.artnr = f.artnr "
    sSQL = sSQL & " set a.Linr = f.linr "
    sSQL = sSQL & " , a.lekpr = f.lekpr  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set linr = 500000 where linr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set lekpr = 0 where lekpr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    
    
    
    'restliche Standards
    sSQL = "Update Artikel set "
    sSQL = sSQL & " gefuehrt = 'J' "
    sSQL = sSQL & " ,BONUS_OK = 'J' "
    sSQL = sSQL & " , RABATT_OK = 'J' "
    sSQL = sSQL & " , UMS_OK = 'J' "
    sSQL = sSQL & " , AWM = '0' "
    sSQL = sSQL & " , PGN = 0 "
    sSQL = sSQL & " , LPZ = 1 "
    sSQL = sSQL & " , RKZ = 'N' "
    sSQL = sSQL & " , PREISSCHU = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    

    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour wird geladen...": lblx.Refresh
    
    
    
    
    'Kassjour
    
    loeschNEW "Bon", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "Bon"
    
    SpalteAnfuegenNEW "Bon", "ADATE", "Datetime", gdBase
    SpalteAnfuegenNEW "Bon", "AZEIT", "Text(8)", gdBase
    
    sSQL = "Create index Zeitpunkt on Bon(Zeitpunkt) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bon set AZEIT = right(Zeitpunkt,8)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bon set ADATE = left(Zeitpunkt,10)"
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "Bonposition", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "Bonposition"
    
    
    

    sSQL = "Create index MWSTsatz on Bonposition(MWSTsatz) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from Kassjour"
    gdBase.Execute sSQL, dbFailOnError
    
    

    
        sSQL = "Insert into Kassjour Select "
        sSQL = sSQL & " artid as Artnr   "
        sSQL = sSQL & ", Bezeichnung as Bezeich  "
        sSQL = sSQL & ", Menge "
        sSQL = sSQL & ", Preis"
        sSQL = sSQL & ", 1 as Filiale   "
        sSQL = sSQL & ", KasseID as Kasnum   "
        sSQL = sSQL & ", BID as BELEGNR  "
        sSQL = sSQL & ", 'BA' as KK_ART   "
        sSQL = sSQL & ", 'J' as UMS_OK   "
        sSQL = sSQL & ", 'V' as MWST   "
        sSQL = sSQL & " from Bonposition  "
        sSQL = sSQL & " where MWSTsatz = 19 "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Kassjour Select "
        sSQL = sSQL & " artid as Artnr   "
        sSQL = sSQL & ", Bezeichnung as Bezeich  "
        sSQL = sSQL & ", Menge "
        sSQL = sSQL & ", Preis"
        sSQL = sSQL & ", 1 as Filiale   "
        sSQL = sSQL & ", KasseID as Kasnum   "
        sSQL = sSQL & ", BID as BELEGNR  "
        sSQL = sSQL & ", 'BA' as KK_ART   "
        sSQL = sSQL & ", 'J' as UMS_OK   "
        sSQL = sSQL & ", 'E' as MWST   "
        sSQL = sSQL & " from Bonposition  "
        sSQL = sSQL & " where MWSTsatz = 7 "
        gdBase.Execute sSQL, dbFailOnError
        

    sSQL = "Create index BID on Bon(BID) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create index KasseID on Bon(KasseID) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Kassjour k inner join Bon b on k.belegnr = b.BID and k.Kasnum = b.KasseID "
    sSQL = sSQL & " set k.adate = b.adate , k.azeit = b.azeit, K.kk_art = b.zahlart"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    sSQL = "Update Kassjour Set KUNDNR = 0 where kundnr is null"
    gdBase.Execute sSQL, dbFailOnError
    

    
    
    
    
    
    
    
    
    loeschNEW "lieferantenartikel", gdBase
    loeschNEW "orgaeinheit", gdBase
    loeschNEW "zeile", gdBase
    loeschNEW "Lieferant", gdBase
    loeschNEW "warengruppe", gdBase
    loeschNEW "bon", gdBase
    loeschNEW "bonposition", gdBase
    loeschNEW "Abschluss", gdBase
    loeschNEW "tblArtikel", gdBase
    loeschNEW "strichcode", gdBase
    loeschNEW "strichcodeTop1", gdBase
    
    
    
    
    
    
    db.Close
    
    
    lblx.Caption = TimeValue(Now) & ": Der Import ist fertig!": lblx.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchapflImport"
        Fehler.gsFehlertext = "Im Programmteil Schapfl Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
    
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    anzeige "normal", "", Label1(4)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Schapfl Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtStatus_Change()
    On Error GoTo LOKAL_ERROR
    
    Dim nProz As Long
  
    nProz = Val(txtstatus.Text)
    ShowProgress picprogress, nProz, 0, 100, True
    picprogress.Refresh

Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtstatus_Change"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub




