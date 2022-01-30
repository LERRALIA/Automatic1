VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL200 
   Caption         =   "Quisy Import"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL200.frx":0000
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
      Caption         =   "Quisy Import"
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
Attribute VB_Name = "frmWKL200"
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
            Unload frmWKL200
        Case 6
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Wo ist die SpielMit - Datei?"

                .Filter = "Access - Dateien (*.mdb)|SpielMit.mdb"
                .ShowSave

                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
            End With
            
            SpielMitImport Label1(4), sPfad
    End Select

err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil SpielMit Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SpielMitImport(lblx As Label, cPfad As String)
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
    
    
    Set db = OpenDatabase(cPfad & "SpielMit.mdb", False, False)
    
    
    
    'Gutscheine
    
    loeschNEW "tblGutschein", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblGutschein"
    
   
    
    sSQL = "Delete * from Gutsch "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into Gutsch Select "
    sSQL = sSQL & " gutnr as gutschnr   "
    sSQL = sSQL & ", ausstell as DAT_AUSG "
    sSQL = sSQL & ", einloes as DAT_EINL "
    sSQL = sSQL & ", kdnr as Kundnr "
    sSQL = sSQL & ", ausbetrag as wert "
    sSQL = sSQL & ", 1 as FILIALE "
    sSQL = sSQL & " from tblGutschein  "
    gdBase.Execute sSQL, dbFailOnError
    
    'Firma
    loeschNEW "tblParameter", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblParameter"
    
   
    
    sSQL = "Delete * from Firma "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into Firma Select "
    sSQL = sSQL & " Name1 as Name   "
    sSQL = sSQL & ", STR as STRASSE "
    sSQL = sSQL & ", PLZORT as PLZ "
    sSQL = sSQL & ", ORT "
    sSQL = sSQL & ", UMSIDENT as STEUERNR "
    
    sSQL = sSQL & " from tblParameter  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    'Filialen
    sSQL = "Delete * from Filialen "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into Filialen Select "
    sSQL = sSQL & " Name1 as FilialName   "
    sSQL = sSQL & ", 1 as FILIALNR "
    sSQL = sSQL & " from tblParameter  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    

    loeschNEW "tblAdressen", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblAdressen"
    
    
    'Lieferanten
    
    sSQL = "Delete from LISRT"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LISRT Select "
    sSQL = sSQL & " Kdnrorg as LINR  "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", plzstr as plz "
    sSQL = sSQL & ", name1 as liefbez "
    sSQL = sSQL & ", liefkurz as  KUERZEL "
    sSQL = sSQL & ", Ort as STADT "
    sSQL = sSQL & ", Telefon as TEL "
    sSQL = sSQL & ", Telefax as FAX "
    sSQL = sSQL & ", kundennr as KUNDNR "
    sSQL = sSQL & ", EMAIL "
    sSQL = sSQL & ", mindbest as awert "

    sSQL = sSQL & " from tblAdressen where Gruppe = '150' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from tblAdressen where Gruppe = '150' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    'Bediener
    
    sSQL = "Delete * from Bedname"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select * from Bedname "
    Set rsrsZ = gdBase.OpenRecordset(sSQL)
    
    ineueBednr = 1
    sSQL = "select * from tblAdressen where Gruppe = '300' or Gruppe = '390' "
    Set rsRSQ = gdBase.OpenRecordset(sSQL)

    If Not rsRSQ.EOF Then
        rsRSQ.MoveFirst
        Do While Not rsRSQ.EOF
            rsrsZ.AddNew
            rsrsZ!BEDNU = ineueBednr
            ineueBednr = ineueBednr + 1
            If Not IsNull(rsRSQ!Name1) Then
                rsrsZ!bedname = rsRSQ!Name1
            End If
            rsrsZ!BEDIENER = 9
            rsrsZ.Update
            rsRSQ.MoveNext
        Loop
    End If

    rsRSQ.Close
    rsrsZ.Close
    
      
    'Kunden
    
    sSQL = "Delete * from Kunden"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Kunden Select "
    sSQL = sSQL & " Kdnrorg as KUNDNR  "
    sSQL = sSQL & ", anrede "
    sSQL = sSQL & ", Ucase(left(anrede, 5)) as kuerzel "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", plzstr as plz "
    sSQL = sSQL & ", name1 as name "
    sSQL = sSQL & ", name2 as notizen "
    sSQL = sSQL & ", Ort as STADT "
    sSQL = sSQL & ", Telefon as TEL "
    sSQL = sSQL & ", Telefax as FAXNR "
    sSQL = sSQL & ", rgrab as Rabatt "
    sSQL = sSQL & ", EMAIL "
    sSQL = sSQL & " from tblAdressen  "
    gdBase.Execute sSQL, dbFailOnError
    
    'Artikelgruppe
    
    loeschNEW "tblProduktgruppen", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblProduktgruppen"
    
    sSQL = "Delete * from AGNDBF"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "select * from AGNDBF "
    Set rsrsZ = gdBase.OpenRecordset(sSQL)
    
    
    sSQL = "select * from tblProduktgruppen  "
    Set rsRSQ = gdBase.OpenRecordset(sSQL)

    If Not rsRSQ.EOF Then
        rsRSQ.MoveFirst
        Do While Not rsRSQ.EOF
            rsrsZ.AddNew
            
            
            If Not IsNull(rsRSQ!PRODGRUP) Then
                If IsNumeric(rsRSQ!PRODGRUP) Then
                    rsrsZ!AGN = rsRSQ!PRODGRUP
                    
                    If Not IsNull(rsRSQ!Text) Then
                        rsrsZ!AGTEXT = Left(rsRSQ!Text, 30)
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
    
    
    'Produktgruppen
    
    loeschNEW "tblProduktgruppen", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblProduktgruppen"
    
    sSQL = "Delete * from AGNDBF"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "select * from AGNDBF "
    Set rsrsZ = gdBase.OpenRecordset(sSQL)
    
    
    sSQL = "select * from tblProduktgruppen  "
    Set rsRSQ = gdBase.OpenRecordset(sSQL)

    If Not rsRSQ.EOF Then
        rsRSQ.MoveFirst
        Do While Not rsRSQ.EOF
            rsrsZ.AddNew
            
            
            If Not IsNull(rsRSQ!PRODGRUP) Then
                If IsNumeric(rsRSQ!PRODGRUP) Then
                    rsrsZ!AGN = rsRSQ!PRODGRUP
                    
                    If Not IsNull(rsRSQ!Text) Then
                        rsrsZ!AGTEXT = Left(rsRSQ!Text, 30)
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
    
    

    
'''    'Produkthauptgruppen
'''
'''    loeschNEW "tblProdukthauptgruppen", gdBase
'''    TransferTab db, cPfad1 & "kissdata.mdb", "tblProdukthauptgruppen"
'''
'''
'''
    loeschNEW "PGNDBF", gdBase


    sSQL = "Create Table PGNDBF ("
    sSQL = sSQL & " PGN integer "
    sSQL = sSQL & " , PGNBEZEICH Text(35) "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
'''
'''
'''
'''
'''
'''    sSQL = "select * from PGNDBF "
'''    Set rsrsZ = gdBase.OpenRecordset(sSQL)
'''
'''
'''    sSQL = "select * from tblProdukthauptgruppen  "
'''    Set rsRSQ = gdBase.OpenRecordset(sSQL)
'''
'''    If Not rsRSQ.EOF Then
'''        rsRSQ.MoveFirst
'''        Do While Not rsRSQ.EOF
'''            rsrsZ.AddNew
'''
'''
'''            If Not IsNull(rsRSQ!PRODHGRP) Then
'''                If IsNumeric(rsRSQ!PRODHGRP) Then
'''                    rsrsZ!PGN = Val(rsRSQ!PRODHGRP)
'''
'''                    If Not IsNull(rsRSQ!Text) Then
'''                        rsrsZ!PGNBEZEICH = Left(rsRSQ!Text, 30)
'''                    Else
'''                        rsrsZ!PGNBEZEICH = ""
'''                    End If
'''
'''
'''                End If
'''            End If
'''
'''
'''
'''
'''            rsrsZ.Update
'''            rsRSQ.MoveNext
'''        Loop
'''    End If
'''
'''    rsRSQ.Close
'''    rsrsZ.Close
'''
'''
'''
'''
'''    sSQL = "Delete * from PGNDBF where pgn = 0 "
'''    gdBase.Execute sSQL, dbFailOnError
'''
'''    sSQL = "Delete * from PGNDBF where pgn is null "
'''    gdBase.Execute sSQL, dbFailOnError
'''
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour wird geladen...": lblx.Refresh
    
    
    
    
    'Kassjour
    
    loeschNEW "tblStatistik", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblStatistik"

    sSQL = "Create index MWST on tblStatistik(MWST) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from Kassjour"
    gdBase.Execute sSQL, dbFailOnError
    
    
    If SpalteInTabellegefundenNEW("Kassjour", "ARTNR_ALT", gdBase) = True Then
        sSQL = "Alter Table Kassjour drop ARTNR_ALT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    SpalteAnfuegenNEW "Kassjour", "ARTNR_ALT", "Text(20)", gdBase
    
    sSQL = "Create index ARTNR_ALT on Kassjour(ARTNR_ALT) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    Dim lJahre As Long
    
    For lJahre = 2003 To 2015
    
        lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour(" & lJahre & " 19%) wird geladen...": lblx.Refresh
    
        sSQL = "Insert into Kassjour Select "
        sSQL = sSQL & " 0 as Artnr   "
        sSQL = sSQL & ", tblStatistik.ARTNR as ARTNR_ALT "
        sSQL = sSQL & ", LIDAT as ADATE "
        sSQL = sSQL & ", '00:00:00' as AZEIT "
        sSQL = sSQL & ", Menge "
        sSQL = sSQL & ", (Menge * (tblStatistik.Preis * 119 / 100)) as Preis"
        sSQL = sSQL & ", 1 as Filiale   "
        sSQL = sSQL & ", 1 as Kasnum   "
        sSQL = sSQL & ", KDNR as KUNDNR  "
'        sSQL = sSQL & ", BELEG as BELEGNR  "
        sSQL = sSQL & ", val(liefnr) as linr  "
        sSQL = sSQL & ", EINDM as EKPR   "
        sSQL = sSQL & ", 'BA' as KK_ART   "
        sSQL = sSQL & ", 'J' as UMS_OK   "
        sSQL = sSQL & ", 'V' as MWST   "
        sSQL = sSQL & " from tblStatistik  "
        sSQL = sSQL & " where MWST = 1 "
        sSQL = sSQL & " and year(lidat) = " & lJahre
        gdBase.Execute sSQL, dbFailOnError
        
        lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour(" & lJahre & " 7%) wird geladen...": lblx.Refresh
        
        sSQL = "Insert into Kassjour Select "
        sSQL = sSQL & " 0 as Artnr   "
        sSQL = sSQL & ", tblStatistik.ARTNR as ARTNR_ALT "
        sSQL = sSQL & ", LIDAT as ADATE "
        sSQL = sSQL & ", '00:00:00' as AZEIT "
        sSQL = sSQL & ", Menge "
        sSQL = sSQL & ", (Menge * (tblStatistik.Preis * 107 / 100)) as Preis"
        sSQL = sSQL & ", 1 as Filiale   "
        sSQL = sSQL & ", 1 as Kasnum   "
        sSQL = sSQL & ", KDNR as KUNDNR  "
'        sSQL = sSQL & ", BELEG as BELEGNR  "
        sSQL = sSQL & ", val(liefnr) as linr  "
        sSQL = sSQL & ", EINDM as EKPR   "
        sSQL = sSQL & ", 'BA' as KK_ART   "
        sSQL = sSQL & ", 'J' as UMS_OK   "
        sSQL = sSQL & ", 'E' as MWST   "
        sSQL = sSQL & " from tblStatistik  "
        sSQL = sSQL & " where MWST = 2 "
        sSQL = sSQL & " and year(lidat) = " & lJahre
        gdBase.Execute sSQL, dbFailOnError
    
    Next lJahre
    
    
    
    
    
    
    
    
    sSQL = "Update Kassjour Set KUNDNR = 0 where kundnr is null"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Umsatz wird geladen...": lblx.Refresh
    
    'Umsatz
    sSQL = "Delete * from Umsatz"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Umsatz Select "
    sSQL = sSQL & " adate as Datum   "
    sSQL = sSQL & " , sum(preis) as umsg1 "
    sSQL = sSQL & " from Kassjour group by adate "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "UMTEMP", gdBase
    
    sSQL = "Select  "
    sSQL = sSQL & " adate as Datum   "
    sSQL = sSQL & " , sum(preis) as umsv1 "
    sSQL = sSQL & " into UMTEMP from Kassjour "
    sSQL = sSQL & " where mwst = 'V'"
    sSQL = sSQL & " group by adate "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz inner join UMTEMP on Umsatz.datum = UMTEMP.datum set Umsatz.umsv1 = UMTEMP.umsv1  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "UMTEMP", gdBase
    
    sSQL = "Select  "
    sSQL = sSQL & " adate as Datum   "
    sSQL = sSQL & " , sum(preis) as umse1 "
    sSQL = sSQL & " into UMTEMP from Kassjour "
    sSQL = sSQL & " where mwst = 'E'"
    sSQL = sSQL & " group by adate "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz inner join UMTEMP on Umsatz.datum = UMTEMP.datum set Umsatz.umse1 = UMTEMP.umse1  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set kunz1 = 0,kred1 = 0, umso1 = 0, ekpr1 = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set umse1 = 0 where umse1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set umsv1 = 0 where umsv1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz Set umsg1 = 0 where umsg1 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artlief wird geladen...": lblx.Refresh
    
    
    
    'Artlief
    
    loeschNEW "tblProdLieferstamm", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblProdLieferstamm"
    
    sSQL = "Delete * from Artlief"
    gdBase.Execute sSQL, dbFailOnError
    
    
    If SpalteInTabellegefundenNEW("Artlief", "ARTNR_ALT", gdBase) = True Then
        sSQL = "Alter Table Artlief drop ARTNR_ALT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    SpalteAnfuegenNEW "Artlief", "ARTNR_ALT", "Text(20)", gdBase
    
    sSQL = "Insert into Artlief Select "
    sSQL = sSQL & " 0 as Artnr   "
    sSQL = sSQL & ", tblProdLieferstamm.ARTNR as ARTNR_ALT "
    sSQL = sSQL & ", LIEFNR as LINR "
    sSQL = sSQL & ", EKPREIS as LEKPR "
    sSQL = sSQL & ", 1 as MINMEN "
    sSQL = sSQL & ", ORGNR as LIBESNR "
    sSQL = sSQL & " from tblProdLieferstamm  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel wird geladen...": lblx.Refresh
    'Artikel
    
    loeschNEW "tblArtikel", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblArtikel"
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel Index erstellen...": lblx.Refresh
    
    sSQL = "Create index auslauf on tblArtikel(auslauf) "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel Löschartikel EX...": lblx.Refresh
    
    sSQL = "Update  tblArtikel set auslauf = 'J' where auslauf = 'L' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update  tblArtikel set auslauf = 'J' where auslauf = 'l' "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Delete * from tblArtikel where auslauf = 'L' "
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Delete * from tblArtikel where auslauf = 'l' "
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from Artikel "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "BTIKEL", gdBase
    
    sSQL = "Select * into BTIKEL from Artikel "
    gdBase.Execute sSQL, dbFailOnError
    
    
    If SpalteInTabellegefundenNEW("BTIKEL", "lfnr", gdBase) = True Then
        sSQL = "Alter Table BTIKEL drop lfnr"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    SpalteAnfuegenNEW "BTIKEL", "lfnr", "autoincrement", gdBase
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Zugang erstellen...": lblx.Refresh
    
    
    'kleine temporäre Zugang
    loeschNEW "tblZugang", gdBase
    
    sSQL = "Create Table tblZugang ( "
    sSQL = sSQL & " Artnr  Long "
    sSQL = sSQL & ", LZLIEFDAT DATETIME "
    sSQL = sSQL & ", LZLIEFNR Long "
    sSQL = sSQL & ", Artnr_ALT Text(20) "
    sSQL = sSQL & " )  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into tblZugang Select "
    sSQL = sSQL & " Artnr as Artnr_ALT  "
    sSQL = sSQL & ", LZLIEFDAT "
    sSQL = sSQL & ", LZLIEFNR "
    sSQL = sSQL & " from tblArtikel "
    gdBase.Execute sSQL, dbFailOnError

    'kleine Ende
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel in Btikel importieren...": lblx.Refresh
    
    
    sSQL = "Insert into BTIKEL Select "
    sSQL = sSQL & " Artnr as Notizen  "
    sSQL = sSQL & ", Artbez as bezeich "
    sSQL = sSQL & ", ean"
    sSQL = sSQL & ", Gruppe as agn "
    sSQL = sSQL & ", 0 as pgn "
    sSQL = sSQL & ", 1 as lpz "
    sSQL = sSQL & ", 'N' as rkz "
    sSQL = sSQL & ", Bestand "
    sSQL = sSQL & ", MWST  "
    sSQL = sSQL & ", VKDM as KVKPR1 "
    sSQL = sSQL & ", VKDM as VKPR "
    sSQL = sSQL & ", EKDM as EKPR "
    sSQL = sSQL & ", HAUPTLIEF as LINR "
    sSQL = sSQL & ", EINFDAT as AUFDAT "
    
    sSQL = sSQL & ", 'J' as RABATT_OK "
    sSQL = sSQL & ", 'J' as GEFUEHRT "
    sSQL = sSQL & ", 'N' as PREISSCHU "
    sSQL = sSQL & ", 'J' as BONUS_OK "
    sSQL = sSQL & ", 'J' as UMS_OK "
    sSQL = sSQL & ", '0' as AWM "
    
    
    sSQL = sSQL & " from tblArtikel "
    sSQL = sSQL & " where ucase(Trim(auslauf)) <> 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel in Btikel importieren Teil 2...": lblx.Refresh
    
    sSQL = "Insert into BTIKEL Select "
    sSQL = sSQL & " Artnr as Notizen  "
    sSQL = sSQL & ", Artbez as bezeich "
    sSQL = sSQL & ", ean"
    sSQL = sSQL & ", Gruppe as agn "
    sSQL = sSQL & ", 0 as pgn "
    sSQL = sSQL & ", 1 as lpz "
    sSQL = sSQL & ", 'J' as rkz "
    sSQL = sSQL & ", Bestand "
    sSQL = sSQL & ", MWST  "
    sSQL = sSQL & ", VKDM as KVKPR1 "
    sSQL = sSQL & ", EKDM as EKPR "
    sSQL = sSQL & ", HAUPTLIEF as LINR "
    sSQL = sSQL & ", EINFDAT as AUFDAT "
    
    sSQL = sSQL & ", 'J' as RABATT_OK "
    sSQL = sSQL & ", 'J' as GEFUEHRT "
    sSQL = sSQL & ", 'N' as PREISSCHU "
    sSQL = sSQL & ", 'J' as BONUS_OK "
    sSQL = sSQL & ", 'J' as UMS_OK "
    sSQL = sSQL & ", '0' as AWM "

    
    sSQL = sSQL & " from tblArtikel "
    sSQL = sSQL & " where ucase(Trim(auslauf)) = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Linbez wird gefüllt...": lblx.Refresh
    
    
    
    sSQL = "Delete * from LINBEZ "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LINBEZ Select "
    sSQL = sSQL & " distinct(hauptlief) as linr "
    sSQL = sSQL & " from tblArtikel "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LINBEZ set LPZ = 1  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LINBEZ inner join lisrt on linbez.linr = lisrt.linr set linbez.LINBEZEICH = lisrt.liefbez  "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel Artnr wird erstellt...": lblx.Refresh
    
    sSQL = "Update BTIKEL set Artnr = 670000 + lfnr "
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("BTIKEL", "lfnr", gdBase) = True Then
        sSQL = "Alter Table BTIKEL drop lfnr"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Insert into artikel select * from btikel"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "BTIKEL", gdBase
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 1...": lblx.Refresh

    sSQL = "Update Artikel set EAN = trim(EAN)"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 2...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = '' where EAN = '0000000000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 3...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,3) where left(EAN,10) = '0000000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 4...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,4) where left(EAN,9) = '000000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 5...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,5) where left(EAN,8) = '00000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 6...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,6) where left(EAN,7) = '0000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 7...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,7) where left(EAN,6) = '000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 8...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,8) where left(EAN,5) = '00000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 9...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,9) where left(EAN,4) = '0000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 10...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,10) where left(EAN,3) = '000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 11...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,11) where left(EAN,2) = '00'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 12...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(EAN,12) where left(EAN,1) = '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 13...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = '' where EAN = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 14...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = '' where EAN = '0 0' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 15...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = '' where EAN = '*none*' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 16...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = '' where right(EAN,1)  = 'A' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 17...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = RTRIM(EAN) "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Behandlung 18...": lblx.Refresh
    
    sSQL = "Update Artikel set EAN = right(ean,len(ean)- 1)  where left(EAN,1)  = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN Duplikate ...": lblx.Refresh
    
    
    Dim lAnzdupli As Long
    Dim rsArt As DAO.Recordset
    Dim cEAN As String
    
    
    'hier ean duplikate rausnehmen
    loeschNEW "alit", gdBase
    sSQL = "select count(ean) as count ,ean into alit from ARTIKEL group by ean having count(ean) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
   
    
    
    
    sSQL = "delete from alit where ean is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from alit where trim(ean) = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    Dim bUpdateEan As Boolean
    
    Set rsrs = gdBase.OpenRecordset("alit", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!EAN) Then
                cEAN = Trim(rsrs!EAN)
            End If
            
            sSQL = "Update Artikel set ean = '' where ean = '" & cEAN & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    
    'Ende Ean Duplikate
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel Mwst E...": lblx.Refresh
    
    
    
    
    
    
    
    sSQL = "Update Artikel set mwst = 'E' where mwst = '2'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel Mwst V...": lblx.Refresh
    
    sSQL = "Update Artikel set mwst = 'V' where mwst <> '2'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean wird geladen...": lblx.Refresh
    
    
    'EAN
    
    loeschNEW "tblArtikelEAN", gdBase
    TransferTab db, cPfad1 & "kissdata.mdb", "tblArtikelEAN"
    
    loeschNEW "ArtikelEAN", gdBase
    
    sSQL = "Create Table ArtikelEAN ( "
    sSQL = sSQL & " Artnr  Long "
    sSQL = sSQL & ", EAN Text(13) "
    sSQL = sSQL & ", Artnr_ALT Text(20) "
    sSQL = sSQL & " )  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into ArtikelEAN Select "
    sSQL = sSQL & " 0 as Artnr   "
    sSQL = sSQL & ",  EAN "
    sSQL = sSQL & ", ARTORGA as Artnr_Alt "
    sSQL = sSQL & " from tblArtikelEAN  "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 1...": lblx.Refresh
    
    
    sSQL = "Update ArtikelEAN set EAN = trim(EAN)"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 2...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = '' where EAN = '0000000000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 3...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,3) where left(EAN,10) = '0000000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 4...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,4) where left(EAN,9) = '000000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 5...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,5) where left(EAN,8) = '00000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 6...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,6) where left(EAN,7) = '0000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 7...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,7) where left(EAN,6) = '000000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 8...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,8) where left(EAN,5) = '00000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 9...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,9) where left(EAN,4) = '0000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 10...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,10) where left(EAN,3) = '000'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 11...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,11) where left(EAN,2) = '00'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 12...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(EAN,12) where left(EAN,1) = '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 13...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = '' where EAN = '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 14...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = '' where EAN = '0 0' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 15...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = '' where EAN = '*none*' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 16...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = '' where right(EAN,1)  = 'A' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 17...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = RTRIM(EAN) "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 18...": lblx.Refresh
    
    sSQL = "Update ArtikelEAN set EAN = right(ean,len(ean)- 1)  where left(EAN,1)  = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 19...": lblx.Refresh
    
    
    sSQL = "Update ArtikelEAN set EAN = '' where EAN is null "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artean EAN Behandlung 20...": lblx.Refresh
    

    sSQL = "Delete from ArtikelEAN where EAN = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel Artnr Abgleich...": lblx.Refresh
    
    
    'Artlief brauch noch die neuen Artikelnummern
    sSQL = "Update Artikel a inner join Artlief b on a.notizen = b.artnr_alt  set b.artnr = a.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle ArtikelEAN Artnr Abgleich...": lblx.Refresh
    
    sSQL = "Update Artikel a inner join ArtikelEAN b on a.notizen = b.artnr_alt  set b.artnr = a.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour Artnr Abgleich...": lblx.Refresh
    
    For lJahre = 2003 To 2015
    
        lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour(" & lJahre & "), Artnr Abgleich...": lblx.Refresh
        
        sSQL = "Update Kassjour a inner join Artikel b on a.artnr_alt = b.notizen set a.artnr = b.artnr "
        sSQL = sSQL & " where year(a.adate) = " & lJahre
        gdBase.Execute sSQL, dbFailOnError
    Next lJahre
    
    lblx.Caption = TimeValue(Now) & ": Tabelle tblZugang Artnr Abgleich...": lblx.Refresh
    
    sSQL = "Update tblZugang a inner join Artikel b on a.artnr_alt = b.notizen set a.artnr = b.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from tblZugang  where artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Zugang füllen...": lblx.Refresh
    
    sSQL = "Insert into ZUGANG Select  "
    sSQL = sSQL & " Artnr "
    sSQL = sSQL & ", LZLIEFDAT  as ADATE"
    sSQL = sSQL & ", LZLIEFNR  as LINR"
    sSQL = sSQL & ", '00:00:00' as UHRZEIT "
    sSQL = sSQL & ", 1 as Bewegung "
    sSQL = sSQL & ", 0 as BESTANDALT "
    sSQL = sSQL & ", 0 as BESTANDNEU "
    sSQL = sSQL & ", 1 as FILIALNR "
    sSQL = sSQL & " from tblZugang "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Zugang Bezeich Abgleich...": lblx.Refresh
    
    sSQL = "Update Zugang a inner join Artikel b on a.artnr = b.artnr set a.bezeich = b.bezeich  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    For lJahre = 2003 To 2015
    
        lblx.Caption = TimeValue(Now) & ": Tabelle Kassjour(" & lJahre & "), Update...": lblx.Refresh
    
        sSQL = "Update Kassjour a inner join Artikel b on a.artnr = b.artnr set a.bezeich = b.bezeich  "
        sSQL = sSQL & ",a.ean = b.ean"
        sSQL = sSQL & ",a.agn = b.agn"
        sSQL = sSQL & ",a.lpz = b.lpz"
        sSQL = sSQL & ",a.vkpr = b.vkpr"
        sSQL = sSQL & " where year(adate) = " & lJahre
        gdBase.Execute sSQL, dbFailOnError
        
    Next lJahre
    
    
    
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle ArtikelEAN EAN Bereinigung...": lblx.Refresh
    
    
    
    sSQL = "Delete * from ArtikelEAN  where EAN In (Select EAN from artikel where EAN <> '') "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = TimeValue(Now) & ": Tabelle ArtikelEAN EAN Duplikate...": lblx.Refresh
    
    
    'hier ean duplikate rausnehmen
    loeschNEW "alit", gdBase
    sSQL = "select count(ean) as count ,ean into alit from ArtikelEAN group by ean having count(ean) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from alit where ean is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from alit where trim(ean) = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("alit", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!EAN) Then
                cEAN = Trim(rsrs!EAN)
            End If
            
            sSQL = "Delete * from ArtikelEAN where ean = '" & cEAN & "'"
            gdBase.Execute sSQL, dbFailOnError
            
            
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    
    
    
    
    
    
    
    
    
'    sSQL = "Update Artikel a inner join ArtikelEAN b on a.artnr = b.artnr  set b.artnr = a.artnr "
'    gdBase.Execute sSQL, dbFailOnError


    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN2 füllen...": lblx.Refresh

    'EAN2 füllen

    
    loeschNEW "MaxEAN", gdBase
    sSQL = "Select Max(ean)as EAN2, artnr into MaxEAN from ArtikelEAN  group by artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Artikel a inner join MaxEAN b on a.artnr = b.artnr  set a.ean2 = b.ean2 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Delete * from ArtikelEAN  where EAN In (Select EAN2 from MaxEAN where EAN2 <> '') "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Artikel EAN3 füllen...": lblx.Refresh
    
    'EAN3 füllen
    
    loeschNEW "MaxEAN", gdBase
    sSQL = "Select Max(ean) as EAN3, artnr into MaxEAN from ArtikelEAN  group by artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update Artikel a inner join MaxEAN b on a.artnr = b.artnr  set a.ean3 = b.ean3 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Delete * from ArtikelEAN  where EAN In (Select EAN3 from MaxEAN where EAN3 <> '') "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle Spalten löschen...": lblx.Refresh
    
    
    If SpalteInTabellegefundenNEW("Kassjour", "ARTNR_ALT", gdBase) = True Then
    
        sSQL = "Drop index ARTNR_ALT on Kassjour "
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "Alter Table Kassjour drop ARTNR_ALT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Artlief", "ARTNR_ALT", gdBase) = True Then
        sSQL = "Alter Table Artlief drop ARTNR_ALT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    If SpalteInTabellegefundenNEW("Artikel", "lfnr", gdBase) = True Then
        sSQL = "Alter Table Artikel drop lfnr"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    lblx.Caption = TimeValue(Now) & ": Tabelle ARTEAN_K  füllen...": lblx.Refresh
    
    sSQL = "Delete * from ARTEAN_K"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ARTEAN_K Select "
    sSQL = sSQL & " Artnr   "
    sSQL = sSQL & ", EAN "
    sSQL = sSQL & " from ArtikelEAN  "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = TimeValue(Now) & ": Tabelle ARTEAN_K EAN Behandlung...": lblx.Refresh
    
    sSQL = "Update ARTEAN_K Set EAN = '0' & EAN where LEN(EAN) = 11"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel Set EAN = '0' & EAN where LEN(EAN) = 11"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel Set EAN2 = '0' & EAN2 where LEN(EAN2) = 11"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel Set EAN3 = '0' & EAN3 where LEN(EAN3) = 11"
    gdBase.Execute sSQL, dbFailOnError
    
    
    db.Close
    
    
    lblx.Caption = TimeValue(Now) & ": Der Quisy-Import ist fertig!": lblx.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SpielMitImport"
        Fehler.gsFehlertext = "Im Programmteil Quisy Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
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
    Fehler.gsFehlertext = "Im Programmteil Quisy Import ist ein Fehler aufgetreten."
    
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
  
    nProz = Val(txtStatus.Text)
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



