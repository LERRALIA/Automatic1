VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL152 
   Caption         =   "Futura Import"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL152.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      Picture         =   "frmWKL152.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   1440
      TabIndex        =   10
      Top             =   240
      Width           =   1440
   End
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   9600
      TabIndex        =   11
      Top             =   6360
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
      Caption         =   "Datapos imp"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "- die Futura Daten (Ordner 'Daten' im Verzeichnis 'Sven')"
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
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   10095
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "- eine leere Winkiss Datenbank (KissWk.mdb)"
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
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   10095
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "- aktuelle KISS Stammdaten (Artikel.dbf)"
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
      Caption         =   "Futura Import"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmWKL152"
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
            Unload frmWKL152
        Case 1
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .FileName = ""
                .DialogTitle = "Wo ist die dataposc.mdb?"

                .Filter = "Access - Dateien (*.mdb)|dataposc.mdb"
                .ShowSave

                sdbPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
            End With
            
            DataPos_Import Label1(4), sdbPfad
        Case 6
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Wo sind die Futura - Dateien?"

                .Filter = "Paradox - Dateien (*.db)|Lager.db"
                .ShowSave

                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
            End With
            
            FuturaImport Label1(4), sPfad
            
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .FileName = ""
                .DialogTitle = "Wo sind die KISS - Stammdaten?"

                .Filter = "dBase - Dateien (*.dbf)|Artikel.dbf"
                .ShowSave

                sdbPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
            End With
            
            FuturaImport2 Label1(4), sPfad, sdbPfad
            FuturaImport3 Label1(4), sPfad
        
    End Select

err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Futura Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub DataPos_Import(lblx As Label, cDataPfad As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    
    Dim dbDataPos           As Database
    
    Dim cPfad       As String
    Dim cOldpath    As String
    Dim cNewpath    As String
    Dim lRet        As Long
    Dim lfail       As Long
    Dim j           As Integer
    

    

    Screen.MousePointer = 11
    
    lblx.Caption = "Winkiss Datenbank wird erstellt..."
    lblx.Refresh
    
    Set dbDataPos = OpenDatabase(cDataPfad & "\dataposc.MDB", False, False, "MS Access;PWD=dpc-zauber")
    
    
    
    
    
    lblx.Caption = "Kassjour wird vorbereitet..."
    lblx.Refresh
    
    
    loeschNEW "REPOS", gdBase
    TransferTab dbDataPos, gcDBPfad & "\Kissdata.mdb", "REPOS"
    
    sSQL = "Delete from Kassjour "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table Kassjour add column Text1 TEXT(10)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Repos set reposArtgrp = 0 where reposArtgrp is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Repos set reposart = 0 where reposart is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Repos set reposArtbez = '' where reposArtbez is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Repos set reposart = val(reposart) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Repos where reposart = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Repos set reposart = '0' & reposart  where len(reposart) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Repos set mwst = 0 where mwst is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Repos set mwst = val(mwst) "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = "Kassjour wird importiert..."
    lblx.Refresh
    
    sSQL = "Insert into Kassjour Select  "
    sSQL = sSQL & " 999999 as artnr "
    sSQL = sSQL & " , reposart as ean  "
    sSQL = sSQL & " , reposSTK as Menge  "
    sSQL = sSQL & " , reposGes as Preis  "
    sSQL = sSQL & " , reposArtbez as Bezeich  "
    sSQL = sSQL & " , val(reposArtgrp) as AGN  "
    
    sSQL = sSQL & " , reposdat as adate  "
    sSQL = sSQL & " , repostime as azeit  "
    sSQL = sSQL & " , val(reposkasse) as kasnum  "
    
    sSQL = sSQL & " , mwst as text1  "
    
    sSQL = sSQL & " , ekpreis as ekpr  "
    
    sSQL = sSQL & " , 0 as Filiale "
    sSQL = sSQL & " , 'J' as UMS_OK "
    
    sSQL = sSQL & " from REPOS "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = "Kassjour MwSt..."
    lblx.Refresh
    
    sSQL = "Update Kassjour set MWST = 'V'   where Text1 = '19' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour set MWST = 'E'   where Text1 = '7' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour set MWST = 'V'   where mwst is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour set MWST = 'V'   where mwst = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    
        txtStatus.Text = 8


    sSQL = "Alter Table Kassjour drop Text1 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    
    

    
    'Kunden
    
    txtStatus.Text = 3
    
    lblx.Caption = "Bediener werden importiert..."
    lblx.Refresh
    
    loeschNEW "MITARB", gdBase
    TransferTab dbDataPos, gcDBPfad & "\Kissdata.mdb", "MITARB"

    txtStatus.Text = 5
    
    sSQL = "Delete from Bedname"
    gdBase.Execute sSQL, dbFailOnError


    sSQL = "insert into Bedname Select  "
    sSQL = sSQL & " Personalnr as Bednu "
    sSQL = sSQL & ", vorname & space(1) &  Name as bedname"
    sSQL = sSQL & ", 9 as bediener"
    sSQL = sSQL & ", 'kiss' as passwort"
    sSQL = sSQL & " from mitarb "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "MITARB", gdBase

    lblx.Caption = "Artikelgruppen werden importiert..."
    lblx.Refresh
        
    loeschNEW "GRUPPE", gdBase
    TransferTab dbDataPos, gcDBPfad & "\Kissdata.mdb", "GRUPPE"
 
    sSQL = "Delete from AGNDBF"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "insert into AGNDBF Select  "
    sSQL = sSQL & " val(grpkz) as AGN "
    sSQL = sSQL & ", grpbez as AGText"
    sSQL = sSQL & " from GRUPPE "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "GRUPPE", gdBase
    
    lblx.Caption = "Lieferanten werden importiert..."
    lblx.Refresh
    
    loeschNEW "LIEFERANT", gdBase
    TransferTab dbDataPos, gcDBPfad & "\Kissdata.mdb", "LIEFERANT"
 
    sSQL = "Delete from LISRT"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table LISRT add column LFNR autoincrement"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "insert into LISRT Select  "
    sSQL = sSQL & " liefkz as KUERZEL "
    sSQL = sSQL & ", lieferant & space(1) & name2 as liefbez"
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", plz "
    sSQL = sSQL & ", Ort as STADT "
    sSQL = sSQL & ", Telefon as TEL "
    sSQL = sSQL & ", Telefax as FAX "
    sSQL = sSQL & ", knr as KUNDNR "
    sSQL = sSQL & ", EMAIL "
    sSQL = sSQL & ", bem as notiz "
    sSQL = sSQL & ", anpartner as ktext "
    sSQL = sSQL & ", handy as adress "
    sSQL = sSQL & " from LIEFERANT "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE LISRT set linr = lfnr + 500010 "
    gdBase.Execute sSQL, dbFailOnError
    
     sSQL = "Alter Table LISRT drop LFNR "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "LIEFERANT", gdBase
    
    sSQL = "Insert into LISRT (Kuerzel,Liefbez,Linr) values ('DUMMY','DUMMY',500000)"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
'    Umsatz

    lblx.Caption = "Umsätze werden importiert..."
    lblx.Refresh

    loeschNEW "BESTAND", gdBase
    TransferTab dbDataPos, gcDBPfad & "\Kissdata.mdb", "BESTAND"
 
    sSQL = "Delete from Umsatz"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "insert into Umsatz Select  "
    sSQL = sSQL & " lud as Datum "
    sSQL = sSQL & ", sum(bestand) as UMSG1"
'
    sSQL = sSQL & " from BESTAND group by lud "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "BESTAND", gdBase
    
    sSQL = "Delete from Umsatz where umsg1 = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set UMSV1 = UMSg1 "
    sSQL = sSQL & ", UMSE1 = 0 "
    sSQL = sSQL & ", UMSO1 = 0 "
    sSQL = sSQL & ", KRED1 = 0 "
    sSQL = sSQL & ", EKPR1 = 0 "
    sSQL = sSQL & ", KUNZ1 = 0 "
    gdBase.Execute sSQL, dbFailOnError

        

    lblx.Caption = "Artikel werden vorbereitet..."
    lblx.Refresh
        
        
    'Artikel
    
    loeschNEW "ArtikelDATA", dbDataPos
    
    sSQL = "Select * into ArtikelData from Artikel  "
    dbDataPos.Execute sSQL, dbFailOnError
    
    loeschNEW "ArtikelDATA", gdBase
    TransferTab dbDataPos, gcDBPfad & "\Kissdata.mdb", "ArtikelDATA"
 
    sSQL = "Delete from Artikel"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table Artikel add column LFNR autoincrement"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE ArtikelDATA set aktbest = 0 where aktbest is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE ArtikelDATA set artgruppe = 0 where artgruppe is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE ArtikelDATA set minbestmenge = 0 where minbestmenge is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE ArtikelDATA set Barcodenr = 0 where Barcodenr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE ArtikelDATA set bar1 = 0 where bar1 is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE ArtikelDATA set bar2 = 0 where bar2 is null "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = "Artikel werden importiert..."
    lblx.Refresh
    
    
    sSQL = "insert into Artikel Select  "
    sSQL = sSQL & " 0 as Artnr "
    sSQL = sSQL & " ,artbez as BEZEICH "
    sSQL = sSQL & " ,vkpreis as kvkpr1 "
    sSQL = sSQL & " ,bestellnr as libesnr "
    sSQL = sSQL & " ,val(aktbest) as bestand "
    sSQL = sSQL & " ,val(artgruppe) as agn "
    sSQL = sSQL & " ,artmwst as mwst "
    sSQL = sSQL & " ,me as inhaltbez "
    sSQL = sSQL & " ,lieferant as notizen "
    sSQL = sSQL & " ,val(minbestmenge) as minmen "
    sSQL = sSQL & " ,ekpreis as lekpr "
    sSQL = sSQL & " ,val(Barcodenr) as ean "
    sSQL = sSQL & " ,val(bar1) as ean2 "
    sSQL = sSQL & " ,val(bar2) as ean3 "
    sSQL = sSQL & " from ArtikelDATA  "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ArtikelDATA", gdBase
    
    lblx.Caption = "Artikel Linr..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel inner join Lisrt on artikel.notizen = lisrt.kuerzel set artikel.LINR= lisrt.linr "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = "Artikel EAN..."
    lblx.Refresh
    
    
    sSQL = "UpdATE Artikel set ean = '' where ean = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set ean2 = '' where ean2 = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set ean3 = '' where ean3 = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean = '0' & ean  where len(ean) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean2 = '0' & ean2  where len(ean2) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean3 = '0' & ean3  where len(ean3) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = "Artikel MwSt..."
    lblx.Refresh
    
    
    sSQL = "UpdATE Artikel set MWST = 'E' where MWST = '2' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set MWST = 'V' where MWST = '1' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set LINR = 500000 where linr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
   
    
    
    
    sSQL = "UpdATE Artikel set Artnr = lfnr + 600000 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table Artikel drop LFNR "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    'Artlief
    
    lblx.Caption = "Artlief wird erstellt..."
    lblx.Refresh
    
    
    sSQL = "Delete from Artlief "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into Artlief Select  "
    sSQL = sSQL & " Artnr "
    sSQL = sSQL & " , Linr "
    sSQL = sSQL & " , libesnr "
    sSQL = sSQL & " , minmen "
    sSQL = sSQL & " , lekpr "
    sSQL = sSQL & " from Artikel  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE Artikel set  "
    sSQL = sSQL & " gefuehrt = 'J' "
    sSQL = sSQL & " , BONUS_OK = 'J' "
    sSQL = sSQL & " , RABATT_OK = 'J' "
    sSQL = sSQL & " , UMS_OK = 'J' "
    sSQL = sSQL & " , AWM = '0' "
    sSQL = sSQL & " , PREISSCHU = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artlief set  "
    sSQL = sSQL & " RKZ = 'N' "
    gdBase.Execute sSQL, dbFailOnError
    
    


    
    dbDataPos.Close
    
    lblx.Caption = TimeValue(Now) & ": DataPos Import ist fertig!": lblx.Refresh
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 3380 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DataPos_Import"
        Fehler.gsFehlertext = "Beim Futura Import Teil 2 ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
    
End Sub
Public Sub FuturaImport(lblx As Label, cPfad As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim dbFUT           As Database
    Dim dbQ             As Database
    Dim lAnzTable       As Long
    Dim lcount          As Long
    Dim sTabname        As String
    Dim lZaehler        As Long
    Dim sTabArray(453)      As String
    Dim i As Integer
    Dim bnoteinlese  As Boolean
    
    
    lblx.Caption = TimeValue(Now) & ": Tabellen werden in den Speicher geladen...": lblx.Refresh
    i = 0
    
    sTabArray(i) = "ART_HIST": i = i + 1
    sTabArray(i) = "ACODEDEF": i = i + 1
    sTabArray(i) = "ACODENUM": i = i + 1
    sTabArray(i) = "ADRTEXT": i = i + 1
    sTabArray(i) = "AEDITDET": i = i + 1
    sTabArray(i) = "AEDITKPF": i = i + 1
    sTabArray(i) = "AIRPORT": i = i + 1
    sTabArray(i) = "AIRLINE": i = i + 1
    sTabArray(i) = "AIRFLIGH": i = i + 1
    sTabArray(i) = "AKT_HOUR": i = i + 1
    sTabArray(i) = "AKT_Kopf": i = i + 1
    sTabArray(i) = "ANG_ADR": i = i + 1
    sTabArray(i) = "ANGEZEIL": i = i + 1
    sTabArray(i) = "ANGHEAD": i = i + 1
    sTabArray(i) = "ANGHIST": i = i + 1
    sTabArray(i) = "ANGPERI": i = i + 1
    sTabArray(i) = "ANGZAHL": i = i + 1
    sTabArray(i) = "ANGZEIL": i = i + 1
    sTabArray(i) = "ANREDEN": i = i + 1
    sTabArray(i) = "ANSITRNS": i = i + 1
    sTabArray(i) = "ART_CODE": i = i + 1
    sTabArray(i) = "ART_LEVE": i = i + 1
    sTabArray(i) = "ART_LFID": i = i + 1
    sTabArray(i) = "ART_LINF": i = i + 1
    sTabArray(i) = "ART_LOTS": i = i + 1
    sTabArray(i) = "ART_STCK": i = i + 1
    sTabArray(i) = "ART_TEXT": i = i + 1
    sTabArray(i) = "ART_USER": i = i + 1
    sTabArray(i) = "ART_VERP": i = i + 1
    sTabArray(i) = "ASRT_FIL": i = i + 1
    sTabArray(i) = "ASRT_KPF": i = i + 1
    sTabArray(i) = "ASRT_ART": i = i + 1
    sTabArray(i) = "AUFHIST": i = i + 1
    sTabArray(i) = "AUFTEXT": i = i + 1
    sTabArray(i) = "AUTRSYS": i = i + 1
    sTabArray(i) = "AUSLAGER": i = i + 1
    sTabArray(i) = "AUTBAUM": i = i + 1
    sTabArray(i) = "AUTTRANS": i = i + 1
    sTabArray(i) = "AUTTEXT": i = i + 1
    sTabArray(i) = "AUTMODUL": i = i + 1
    'B
    sTabArray(i) = "BANKNAME": i = i + 1
    sTabArray(i) = "BASART": i = i + 1
    sTabArray(i) = "BASHEAD": i = i + 1
    sTabArray(i) = "BAT_HEAD": i = i + 1
    sTabArray(i) = "BAT_PROT": i = i + 1
    sTabArray(i) = "BATCHDET": i = i + 1
    sTabArray(i) = "BATCHGRP": i = i + 1
    sTabArray(i) = "BATCHJOB": i = i + 1
    sTabArray(i) = "BATCHLIN": i = i + 1
    sTabArray(i) = "BATCHMOD": i = i + 1
    sTabArray(i) = "BATCHQUE": i = i + 1
    sTabArray(i) = "BEDIHEAD": i = i + 1
    sTabArray(i) = "BEDING": i = i + 1
    sTabArray(i) = "BEDIZEIL": i = i + 1
    sTabArray(i) = "BENGMEMB": i = i + 1
    sTabArray(i) = "BENGROUP": i = i + 1
    sTabArray(i) = "BENUTZER": i = i + 1
    sTabArray(i) = "BEST_ADR": i = i + 1
    sTabArray(i) = "BESTCOD1": i = i + 1
    sTabArray(i) = "BESTCOD2": i = i + 1
    sTabArray(i) = "BESTHEAD": i = i + 1
    sTabArray(i) = "BESTHIST": i = i + 1
    sTabArray(i) = "BESTLINK": i = i + 1
    sTabArray(i) = "BESTLT": i = i + 1
    sTabArray(i) = "BESTPROZ": i = i + 1
    sTabArray(i) = "BESTRAHM": i = i + 1
    sTabArray(i) = "BESTVOR": i = i + 1
    sTabArray(i) = "BESTWUN": i = i + 1
    sTabArray(i) = "BESTZEIL": i = i + 1
    sTabArray(i) = "BONUS": i = i + 1
    sTabArray(i) = "BSTAINLT": i = i + 1
    sTabArray(i) = "BSTATXLT": i = i + 1
    sTabArray(i) = "BSTBLKLT": i = i + 1
    sTabArray(i) = "BSTBONLT": i = i + 1
    sTabArray(i) = "BSTCODLT": i = i + 1
    sTabArray(i) = "BSTHDRLT": i = i + 1
    sTabArray(i) = "BSTPOSLT": i = i + 1
    sTabArray(i) = "BSTPRSLT": i = i + 1
    sTabArray(i) = "BSTPRZLT": i = i + 1
    sTabArray(i) = "BSTTXTLT": i = i + 1
    sTabArray(i) = "BUDGETD": i = i + 1
    sTabArray(i) = "BUDGETND": i = i + 1
    sTabArray(i) = "BUDGETDF": i = i + 1
    sTabArray(i) = "BUDGETFK": i = i + 1
    sTabArray(i) = "BUDGETZD": i = i + 1
    sTabArray(i) = "BUDGETZL": i = i + 1
    sTabArray(i) = "BUDGORD": i = i + 1
    'C
    sTabArray(i) = "CASH": i = i + 1
    sTabArray(i) = "CASHUSER": i = i + 1
    sTabArray(i) = "CCD_TEST": i = i + 1
    sTabArray(i) = "CCD_INFO": i = i + 1
    sTabArray(i) = "CCD_KOPF": i = i + 1
    sTabArray(i) = "CONFEINT": i = i + 1
    sTabArray(i) = "CONFENUM": i = i + 1
    sTabArray(i) = "CONFGRUP": i = i + 1
    sTabArray(i) = "CONFHILF": i = i + 1
    sTabArray(i) = "CONFRECH": i = i + 1
    sTabArray(i) = "CONFWERT": i = i + 1
    sTabArray(i) = "CUPDHILF": i = i + 1
    sTabArray(i) = "CUPDGRUP": i = i + 1
    sTabArray(i) = "CUPDENUM": i = i + 1
    sTabArray(i) = "CUPDEINT": i = i + 1
    'D
    sTabArray(i) = "DANG_ADR": i = i + 1
    sTabArray(i) = "DANGEZEI": i = i + 1
    sTabArray(i) = "DANGHEAD": i = i + 1
    sTabArray(i) = "DANGHIST": i = i + 1
    sTabArray(i) = "DANGPERI": i = i + 1
    sTabArray(i) = "DANGZAHL": i = i + 1
    sTabArray(i) = "DANGZEIL": i = i + 1
    sTabArray(i) = "DELINFO": i = i + 1
    sTabArray(i) = "DEP_ART": i = i + 1
    sTabArray(i) = "DEP_KOPF": i = i + 1
    sTabArray(i) = "DEP_ZAHL": i = i + 1
    sTabArray(i) = "DEPOTART": i = i + 1
    sTabArray(i) = "DIS_BEDI": i = i + 1
    sTabArray(i) = "DIS_EXTR": i = i + 1
    sTabArray(i) = "DIS_KOPF": i = i + 1
    sTabArray(i) = "DIS_WGR": i = i + 1
    sTabArray(i) = "DIS_KUND": i = i + 1
    sTabArray(i) = "DIZ_ZUS": i = i + 1
    sTabArray(i) = "DLIF_ADR": i = i + 1
    sTabArray(i) = "DLIFHEAD": i = i + 1
    sTabArray(i) = "DLIFHIST": i = i + 1
    sTabArray(i) = "DLIFZAHL": i = i + 1
    sTabArray(i) = "DLIFZEIL": i = i + 1
    sTabArray(i) = "DOCCNTHD": i = i + 1
    sTabArray(i) = "DOCCNTRF": i = i + 1
    sTabArray(i) = "DOCMAPLN": i = i + 1
    sTabArray(i) = "DOCMAPHD": i = i + 1
    sTabArray(i) = "DR_ZUORD": i = i + 1
    sTabArray(i) = "DRECNUNG": i = i + 1
    sTabArray(i) = "DRUCKER": i = i + 1
    sTabArray(i) = "DWEXPORT": i = i + 1
    sTabArray(i) = "DRECZAHL": i = i + 1
    
    'E
    sTabArray(i) = "EANVWKPF": i = i + 1
    sTabArray(i) = "EANVWLST": i = i + 1
    sTabArray(i) = "EC_SPERR": i = i + 1
    sTabArray(i) = "EDI_IDAT": i = i + 1
    sTabArray(i) = "EDI_IN": i = i + 1
    sTabArray(i) = "EDI_ODAT": i = i + 1
    sTabArray(i) = "EDI_OUT": i = i + 1
    sTabArray(i) = "EDISYS": i = i + 1
    sTabArray(i) = "EFTPROTO": i = i + 1
    sTabArray(i) = "EIGEN": i = i + 1
    sTabArray(i) = "EIGLIST": i = i + 1
    sTabArray(i) = "EINLIST": i = i + 1
    sTabArray(i) = "EKATABON": i = i + 1
    sTabArray(i) = "EKATACOD": i = i + 1
    sTabArray(i) = "EKATALOG": i = i + 1
    sTabArray(i) = "EKATARTI": i = i + 1
    sTabArray(i) = "EKATATXT": i = i + 1
    sTabArray(i) = "EKATAVAR": i = i + 1
    sTabArray(i) = "EKATVPRS": i = i + 1
    sTabArray(i) = "EM_DATEN": i = i + 1
    sTabArray(i) = "EM_KOPF": i = i + 1
    sTabArray(i) = "EPARTFIL": i = i + 1
    sTabArray(i) = "EPARTKEN": i = i + 1
    sTabArray(i) = "EPARTMAP": i = i + 1
    sTabArray(i) = "EPARTNER": i = i + 1
    sTabArray(i) = "EPMSGHST": i = i + 1
    sTabArray(i) = "EPRECST": i = i + 1
    sTabArray(i) = "EREPDATA": i = i + 1
    sTabArray(i) = "EREPKOPF": i = i + 1
    sTabArray(i) = "EREPPARA": i = i + 1
    sTabArray(i) = "EREPUSER": i = i + 1
    sTabArray(i) = "EX_FDATA": i = i + 1
    sTabArray(i) = "EX_FKOPF": i = i + 1
    sTabArray(i) = "EX_LADR": i = i + 1
    sTabArray(i) = "EX_LHEAD": i = i + 1
    sTabArray(i) = "EX_LHIST": i = i + 1
    sTabArray(i) = "EX_LZAHL": i = i + 1
    sTabArray(i) = "EX_LZEIL": i = i + 1
    sTabArray(i) = "EX_RECH": i = i + 1
    sTabArray(i) = "EX_RZAHL": i = i + 1
    sTabArray(i) = "EX_WE_LH": i = i + 1
    sTabArray(i) = "EX_WE_LN": i = i + 1
    sTabArray(i) = "EX_WE_LZ": i = i + 1
    sTabArray(i) = "EX_WE_RH": i = i + 1
    sTabArray(i) = "EX_WE_RN": i = i + 1
    sTabArray(i) = "EX_WELNK": i = i + 1
    sTabArray(i) = "EX_WLHST": i = i + 1
    sTabArray(i) = "EX_WRHST": i = i + 1
    sTabArray(i) = "EXA_CODE": i = i + 1
    sTabArray(i) = "EXA_EANS": i = i + 1
    sTabArray(i) = "EXA_KOPF": i = i + 1
    sTabArray(i) = "EXA_PALG": i = i + 1
    sTabArray(i) = "EXA_PHST": i = i + 1
    sTabArray(i) = "EXA_PRGR": i = i + 1
    sTabArray(i) = "EXARTIKEL": i = i + 1
    sTabArray(i) = "EXPO_HDR": i = i + 1
    sTabArray(i) = "EXPO_FLD": i = i + 1
    sTabArray(i) = "EXPO_DFL": i = i + 1
    sTabArray(i) = "EXTRAFEE": i = i + 1
    sTabArray(i) = "EXTWECHS": i = i + 1
    
    'F
    sTabArray(i) = "FBUINFO": i = i + 1
    sTabArray(i) = "FBUKOPF": i = i + 1
    sTabArray(i) = "FCODEDEF": i = i + 1
    sTabArray(i) = "FCODENUM": i = i + 1
    sTabArray(i) = "FERKOPF": i = i + 1
    sTabArray(i) = "FIBERF": i = i + 1
    sTabArray(i) = "FIBKASDL": i = i + 1
    sTabArray(i) = "FIBKASOP": i = i + 1
    sTabArray(i) = "FIBKASTR": i = i + 1
    sTabArray(i) = "FIBUBUCH": i = i + 1
    sTabArray(i) = "FIBUBWA": i = i + 1
    sTabArray(i) = "FIBUBWAK": i = i + 1
    sTabArray(i) = "FIBUKTXT": i = i + 1
    sTabArray(i) = "FIBUOP": i = i + 1
    sTabArray(i) = "FIL_RAB": i = i + 1
    sTabArray(i) = "FILPRART": i = i + 1
    sTabArray(i) = "FILPRHDR": i = i + 1
    sTabArray(i) = "FILSYSTM": i = i + 1
    sTabArray(i) = "FS_ADDAT": i = i + 1
    sTabArray(i) = "FS_CTRL": i = i + 1
    sTabArray(i) = "FS_MOVE": i = i + 1
    sTabArray(i) = "FS_PURCH": i = i + 1
    sTabArray(i) = "FS_SALES": i = i + 1
    sTabArray(i) = "FTR_KOPF": i = i + 1
    sTabArray(i) = "FTR_DATA": i = i + 1
    
    'K
    sTabArray(i) = "KASS_IMP": i = i + 1
    'L
    sTabArray(i) = "L_AUS": i = i + 1
    sTabArray(i) = "L_DET": i = i + 1
    sTabArray(i) = "L_GND": i = i + 1
    sTabArray(i) = "L_RES": i = i + 1
    sTabArray(i) = "L_SPR": i = i + 1
    sTabArray(i) = "L_SYS": i = i + 1
    sTabArray(i) = "L_TIT": i = i + 1
    sTabArray(i) = "LAGDELTA": i = i + 1
    sTabArray(i) = "LAGTRANS": i = i + 1
    sTabArray(i) = "Lastschr": i = i + 1
    sTabArray(i) = "LAUFSCHL": i = i + 1
    sTabArray(i) = "LFARBE": i = i + 1
    
    sTabArray(i) = "LGORT": i = i + 1
    sTabArray(i) = "LGORTDEF": i = i + 1
    sTabArray(i) = "LGROESSE": i = i + 1
    sTabArray(i) = "LIEFHIST": i = i + 1
    sTabArray(i) = "LIEFZAHL": i = i + 1
    sTabArray(i) = "LIFFPRIS": i = i + 1
    sTabArray(i) = "LIFKPRIS": i = i + 1
    sTabArray(i) = "LIZENZ": i = i + 1
    sTabArray(i) = "LMAPPING": i = i + 1
    sTabArray(i) = "LOESCHDF": i = i + 1
    sTabArray(i) = "LOGONOFF": i = i + 1
    sTabArray(i) = "LSNAPDET": i = i + 1
    sTabArray(i) = "LSNAPHDR": i = i + 1
    'M
    sTabArray(i) = "MAIL": i = i + 1
    sTabArray(i) = "MAIL_REC": i = i + 1
    sTabArray(i) = "MAIL_SND": i = i + 1
    sTabArray(i) = "MANDANT": i = i + 1
    sTabArray(i) = "MANUCASH": i = i + 1
    sTabArray(i) = "MANUHEAD": i = i + 1
    sTabArray(i) = "METHODEN": i = i + 1
    sTabArray(i) = "MNGTYP": i = i + 1
    'N
    sTabArray(i) = "NETZUSER": i = i + 1
    sTabArray(i) = "NL_PARAM": i = i + 1
    'O
    sTabArray(i) = "ORGABT": i = i + 1
    'P
    sTabArray(i) = "PA_WERTE": i = i + 1
    sTabArray(i) = "PACKUNG": i = i + 1
    sTabArray(i) = "PB_ART": i = i + 1
    sTabArray(i) = "PB_SET": i = i + 1
    sTabArray(i) = "PB_TRNS": i = i + 1
    sTabArray(i) = "PILOTDTA": i = i + 1
    sTabArray(i) = "PILOTINF": i = i + 1
    sTabArray(i) = "PILOTKPF": i = i + 1
    sTabArray(i) = "PILOTTRE": i = i + 1
    sTabArray(i) = "PLZ_REF": i = i + 1
    sTabArray(i) = "PR_DTL": i = i + 1
    sTabArray(i) = "PRAEMDET": i = i + 1
    sTabArray(i) = "PREISGRU": i = i + 1
    sTabArray(i) = "PREISLAG": i = i + 1
    sTabArray(i) = "PREISPKT": i = i + 1
    sTabArray(i) = "PREISRND": i = i + 1
    sTabArray(i) = "PREISSTF": i = i + 1
    sTabArray(i) = "PRINTER": i = i + 1
    sTabArray(i) = "PRINTMAP": i = i + 1
    sTabArray(i) = "PRINTQUE": i = i + 1
    sTabArray(i) = "PRODDET": i = i + 1
    sTabArray(i) = "PRODGRP": i = i + 1
    sTabArray(i) = "PRODHEAD": i = i + 1
    sTabArray(i) = "PRODTEIL": i = i + 1
    sTabArray(i) = "PRODTOUR": i = i + 1
    sTabArray(i) = "PROMOART": i = i + 1
    sTabArray(i) = "PROMOTN": i = i + 1
    sTabArray(i) = "PVERFIL": i = i + 1
    sTabArray(i) = "PVERGRP": i = i + 1
    sTabArray(i) = "PVERZUS": i = i + 1
    
    'R
    sTabArray(i) = "RECHNUNG": i = i + 1
    sTabArray(i) = "RECHZAHL": i = i + 1
    sTabArray(i) = "REGION": i = i + 1
    sTabArray(i) = "RETGRUND": i = i + 1
    sTabArray(i) = "RPT_SOG": i = i + 1
    sTabArray(i) = "RPT_SOK": i = i + 1
    sTabArray(i) = "RPT_VTG": i = i + 1
    sTabArray(i) = "RPT_VTP": i = i + 1
    
    
    
    'G
    sTabArray(i) = "GWVORDET": i = i + 1
    sTabArray(i) = "GWVORSUM": i = i + 1
    'I
    sTabArray(i) = "INVLKGRD": i = i + 1
    
    
    'S
    sTabArray(i) = "SBLAYCOL": i = i + 1
    sTabArray(i) = "SBLAYHDR": i = i + 1
    sTabArray(i) = "SBLAYROW": i = i + 1
    sTabArray(i) = "SCANDATA": i = i + 1
    sTabArray(i) = "SCANKOPF": i = i + 1
    sTabArray(i) = "SERDATA": i = i + 1
    sTabArray(i) = "SERDELTA": i = i + 1
    sTabArray(i) = "SERTRANS": i = i + 1
    sTabArray(i) = "SOMI_DET": i = i + 1
    sTabArray(i) = "SOMI_FIL": i = i + 1
    sTabArray(i) = "SOMI_KPF": i = i + 1
    sTabArray(i) = "SOR_ARTI": i = i + 1
    sTabArray(i) = "SOR_KOPF": i = i + 1
    sTabArray(i) = "SPOOLING": i = i + 1
    sTabArray(i) = "SPRACHE": i = i + 1
    sTabArray(i) = "SSTATDAT": i = i + 1
    sTabArray(i) = "SSTATHDR": i = i + 1
    sTabArray(i) = "STATAFLG": i = i + 1
    sTabArray(i) = "STATART": i = i + 1
    sTabArray(i) = "STATART2": i = i + 1
    sTabArray(i) = "STATDDTA": i = i + 1
    sTabArray(i) = "STATDFLG": i = i + 1
    sTabArray(i) = "STATDIDX": i = i + 1
    sTabArray(i) = "STATKOMP": i = i + 1
    sTabArray(i) = "STATODKY": i = i + 1
    sTabArray(i) = "STATODTA": i = i + 1
    sTabArray(i) = "STATOIDX": i = i + 1
    sTabArray(i) = "STATOIKY": i = i + 1
    sTabArray(i) = "STATPDTA": i = i + 1
    sTabArray(i) = "STATPERI": i = i + 1
    sTabArray(i) = "STATPIDX": i = i + 1
    sTabArray(i) = "STATSORT": i = i + 1
    sTabArray(i) = "STATSYS": i = i + 1
    
    'K
    
    sTabArray(i) = "KARKOPF": i = i + 1
    sTabArray(i) = "KARZEIL": i = i + 1
    sTabArray(i) = "KASCOART": i = i + 1
    sTabArray(i) = "KASCOFIL": i = i + 1
    sTabArray(i) = "KASCOGRP": i = i + 1
    sTabArray(i) = "KASCOKPF": i = i + 1
    sTabArray(i) = "KASFISIA": i = i + 1
    sTabArray(i) = "KASFISID": i = i + 1
    sTabArray(i) = "KASFISIH": i = i + 1
    sTabArray(i) = "KASS_EAN": i = i + 1
    sTabArray(i) = "KAS_IMP": i = i + 1
    sTabArray(i) = "KASS_PLU": i = i + 1
    sTabArray(i) = "KASSDLTA": i = i + 1
    sTabArray(i) = "KASSDUPL": i = i + 1
    sTabArray(i) = "KASSE": i = i + 1
    sTabArray(i) = "KASSWAHL": i = i + 1
    sTabArray(i) = "KB_ADR": i = i + 1
    sTabArray(i) = "KB_DET": i = i + 1
    sTabArray(i) = "KB_KOPF": i = i + 1
    sTabArray(i) = "KLASSIF": i = i + 1
    sTabArray(i) = "KO_DATEN": i = i + 1
    sTabArray(i) = "KO_LINK": i = i + 1
    sTabArray(i) = "KO_SEITE": i = i + 1
    sTabArray(i) = "KO_KOPF": i = i + 1
    sTabArray(i) = "KONDIT": i = i + 1
    sTabArray(i) = "KONTODET": i = i + 1
    sTabArray(i) = "KONTOGRP": i = i + 1
    sTabArray(i) = "KONTOSUM": i = i + 1
    sTabArray(i) = "KOSTENST": i = i + 1
    sTabArray(i) = "KRDAUSGL": i = i + 1
    sTabArray(i) = "KRDDELTA": i = i + 1
    sTabArray(i) = "KRDHIST": i = i + 1
    sTabArray(i) = "KRDKARTE": i = i + 1
    sTabArray(i) = "KRDKAUF": i = i + 1
    sTabArray(i) = "KRDKONTO": i = i + 1
    sTabArray(i) = "KRDZAHL": i = i + 1
    sTabArray(i) = "KST_SPLI": i = i + 1
    sTabArray(i) = "KUN_EIG": i = i + 1
    
   
    'T
    
    sTabArray(i) = "TBSTKOPF": i = i + 1
    sTabArray(i) = "TBSTZEIL": i = i + 1
    sTabArray(i) = "TITELREF": i = i + 1
    sTabArray(i) = "TOUR": i = i + 1
    sTabArray(i) = "TRANSLAT": i = i + 1
    sTabArray(i) = "TRENDLST": i = i + 1
    sTabArray(i) = "TRN_BILD": i = i + 1
    sTabArray(i) = "TRN_MINF": i = i + 1
    
    
    
    
    'U
    sTabArray(i) = "UINFDATA": i = i + 1
    sTabArray(i) = "UINFHEAD": i = i + 1
    sTabArray(i) = "UML_DATA": i = i + 1
    sTabArray(i) = "UML_KOPF": i = i + 1
    sTabArray(i) = "UMSDELTA": i = i + 1
    sTabArray(i) = "UMSTRANS": i = i + 1
    sTabArray(i) = "UPDBAUM": i = i + 1
    sTabArray(i) = "UPDMODUL": i = i + 1
    sTabArray(i) = "UPDTRANS": i = i + 1
    sTabArray(i) = "USEREXEC": i = i + 1
    
    
    
    'V
    sTabArray(i) = "VERGLMGT": i = i + 1
    sTabArray(i) = "VER_KOST": i = i + 1
    sTabArray(i) = "VER_KOPF": i = i + 1
    sTabArray(i) = "VER_BER": i = i + 1
    sTabArray(i) = "VD_ZAHL": i = i + 1
    sTabArray(i) = "VD_LOCK": i = i + 1
    sTabArray(i) = "VD_KLAS": i = i + 1
    sTabArray(i) = "VD_HEAD": i = i + 1
    sTabArray(i) = "VD_FILI": i = i + 1
    sTabArray(i) = "VD_DEFI": i = i + 1
    sTabArray(i) = "VTG_ARTI": i = i + 1
    sTabArray(i) = "VTG_BEDI": i = i + 1
    sTabArray(i) = "VTG_DEBI": i = i + 1
    sTabArray(i) = "VK_BER": i = i + 1
    sTabArray(i) = "VTG_DIST": i = i + 1
    sTabArray(i) = "VTG_HERS": i = i + 1
    sTabArray(i) = "VTG_INFO": i = i + 1
    sTabArray(i) = "VTG_PTAB": i = i + 1
    sTabArray(i) = "VTG_PREI": i = i + 1
    sTabArray(i) = "VTG_KUND": i = i + 1
    sTabArray(i) = "VTG_KOPF": i = i + 1
    'W
    sTabArray(i) = "WAEHRUNG": i = i + 1
    sTabArray(i) = "WARKOPAR": i = i + 1
    sTabArray(i) = "WARKORB": i = i + 1
    sTabArray(i) = "WBL_KOPF": i = i + 1
    sTabArray(i) = "WDEFHEAD": i = i + 1
    sTabArray(i) = "WDEFSELE": i = i + 1
    sTabArray(i) = "WE_HEAD": i = i + 1
    sTabArray(i) = "WE_HEADR": i = i + 1
    sTabArray(i) = "WE_KOST": i = i + 1
    sTabArray(i) = "WE_LHIST": i = i + 1
    sTabArray(i) = "WE_LINK": i = i + 1
    sTabArray(i) = "WE_NEBEN": i = i + 1
    sTabArray(i) = "WE_RECH": i = i + 1
    sTabArray(i) = "WE_RGNEB": i = i + 1
    sTabArray(i) = "WE_RHIST": i = i + 1
    sTabArray(i) = "WE_ZEIL": i = i + 1
    sTabArray(i) = "WFACTDET": i = i + 1
    sTabArray(i) = "WFACTION": i = i + 1
    sTabArray(i) = "WFCOND": i = i + 1
    sTabArray(i) = "WFLGEVNT": i = i + 1
    sTabArray(i) = "WFLGHEAD": i = i + 1
    sTabArray(i) = "WFRULE": i = i + 1
    sTabArray(i) = "WRKBVKI": i = i + 1
    sTabArray(i) = "WRKDELTA": i = i + 1
    sTabArray(i) = "WRKLAGER": i = i + 1
    'Z
    sTabArray(i) = "ZUTEILPR": i = i + 1
    sTabArray(i) = "ZUTEILHD": i = i + 1
    sTabArray(i) = "ZEITZONE": i = i + 1
    sTabArray(i) = "ZEITTYP": i = i + 1
    
    sTabArray(i) = "LAND": i = i + 1
    sTabArray(i) = "KASSKOPF": i = i + 1
    sTabArray(i) = "KAS_EX_K": i = i + 1
    sTabArray(i) = "KASSDKPF": i = i + 1
    sTabArray(i) = "KAS_TKPF": i = i + 1
    sTabArray(i) = "KAS_TEXT": i = i + 1
    sTabArray(i) = "AKT_FIL": i = i + 1
    sTabArray(i) = "AKT_ART": i = i + 1
    sTabArray(i) = "ARBKOPF": i = i + 1
    sTabArray(i) = "KERDEF": i = i + 1
    sTabArray(i) = "KERDISP": i = i + 1
    sTabArray(i) = "MISGRDEF": i = i + 1
    sTabArray(i) = "KASSKUND": i = i + 1
    sTabArray(i) = "KAS_EX_D_OLD": i = i + 1
    sTabArray(i) = "ISETINFO": i = i + 1
    sTabArray(i) = "LAGERKOR": i = i + 1
    sTabArray(i) = "WFLGHEAD_1": i = i + 1
    sTabArray(i) = "KERENTW"
    
    
    picprogress.Visible = True
    lblx.Caption = TimeValue(Now) & ": Datenbank wird erstellt...": lblx.Refresh
    
    Set dbQ = OpenDatabase(cPfad, False, False, "Paradox 5.x;")
    
    Kill cPfad & "\FUTURA.MDB"
    Set dbFUT = CreateDatabase(cPfad & "\FUTURA.MDB", dbLangGeneral, dbVersion40)
    
    lblx.Caption = TimeValue(Now) & ": Datenbank wird aktualisiert...": lblx.Refresh
    dbQ.TableDefs.Refresh
    lAnzTable = dbQ.TableDefs.Count
    lZaehler = lAnzTable
    
    
    
    
    Dim td As TableDef
    Dim siAnzeige As Single
    
    siAnzeige = 0
    For lcount = 0 To lAnzTable - 1
        sTabname = dbQ.TableDefs(lcount).name
        
        siAnzeige = siAnzeige + 1
        txtStatus.Text = CStr((100 * siAnzeige) / 453)
        
        i = 0
        bnoteinlese = False
        For i = 0 To 453
            If UCase(sTabArray(i)) = UCase(sTabname) Then
                bnoteinlese = True
                Exit For
            End If
        Next i
        
        If bnoteinlese = False Then

            Set td = dbFUT.CreateTableDef(sTabname)
            td.Connect = "Paradox 5.x;Database=" & cPfad      'c:\Daten"
            td.SourceTableName = sTabname
            
            
            dbFUT.TableDefs.Append td
            
            lblx.Caption = TimeValue(Now) & ": (" & lZaehler & ") " & sTabname & " wird importiert...": lblx.Refresh

        End If

        lZaehler = lZaehler - 1

    Next lcount
    
    dbFUT.Close
    
    lblx.Caption = TimeValue(Now) & ": Futura Import Teil 1 ist fertig!": lblx.Refresh

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "FuturaImport"
        Fehler.gsFehlertext = "Beim Futura Import Teil 1 ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
    
End Sub
Public Sub FuturaImport2(lblx As Label, cPfadFut As String, cpfaddb As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    Dim dbWK        As Database
    Dim dbFUT       As Database
    Dim cPfad       As String
    Dim cOldpath    As String
    Dim cNewpath    As String
    Dim lRet        As Long
    Dim lfail       As Long
    Dim j           As Integer
    
'    cpfaddb = "C:\DBASE"
    

    Screen.MousePointer = 11
    
    lblx.Caption = "Winkiss Datenbank wird erstellt..."
    lblx.Refresh
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "WKLEER\"
    
    If FileExists(cPfad & "KissWK.mdb") = False Then
        Screen.MousePointer = 0
        lblx.Caption = "Abbruch"
        lblx.Refresh
        Exit Sub
    End If
    
    
    cOldpath = cPfad
    cOldpath = ShortPath(cOldpath)
    cOldpath = cOldpath & "kissWK.mdb"
    
    
    cNewpath = cPfad
    cNewpath = ShortPath(cNewpath)
    cNewpath = cNewpath & "KissFUT.mdb"
    lRet = CopyFile(cOldpath, cNewpath, lfail)
    

    If lRet = 0 Then
        Screen.MousePointer = 0
        lblx.Caption = "Abbruch"
        lblx.Refresh
        Exit Sub
    End If
    
    Set dbFUT = OpenDatabase(cPfadFut & "\FUTURA.MDB", False, False)
    Set dbWK = OpenDatabase(cPfad & "KissFUT.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    'Kunden
    
    txtStatus.Text = 3
    
    lblx.Caption = "Kunden werden exportiert..."
    lblx.Refresh
    
    loeschNEW "ANSCHRIF", dbWK
    TransferTab dbFUT, cPfad & "KissFUT.mdb", "ANSCHRIF"
        
    txtStatus.Text = 5
        
        sSQL = "insert into Kunden Select  "
        sSQL = sSQL & " ans_nummer as KUNDNR "
        sSQL = sSQL & ", ans_name1 as Name"
        sSQL = sSQL & ", ans_name2 as vorname"
        sSQL = sSQL & ", ans_strasse as strasse"
        sSQL = sSQL & ", ans_plz as PLZ"
        sSQL = sSQL & ", ans_ort as stadt"
        sSQL = sSQL & ", ans_anrede as anrede"
        sSQL = sSQL & ", ans_titel as titel"
        sSQL = sSQL & ", ans_telefon as tel"
        sSQL = sSQL & ", ans_telefax as faxnr"
        sSQL = sSQL & ", ans_sachgeburtstag as datum1"
        sSQL = sSQL & ", ans_email as email"
        sSQL = sSQL & " from ANSCHRIF"
        sSQL = sSQL & " where ans_typ = 3 "
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 6
        
        sSQL = "Delete from  Kunden"
        sSQL = sSQL & " where name is null "
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 7
        
        sSQL = "UpdATE Kunden set KUERZEL = UCASE(LEFT(NAME,5))"
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 8
        
        sSQL = "UpdATE Kunden set Anrede = 'Frau'"
        sSQL = sSQL & " where Ucase(LEFT(Anrede,1))= 'F'"
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 9
        
        sSQL = "UpdATE Kunden set Anrede = 'Herr'"
        sSQL = sSQL & " where Ucase(LEFT(Anrede,1))= 'H'"
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 10
        
        sSQL = "UpdATE Kunden set GESCHLECHT = 'W'"
        sSQL = sSQL & " where Ucase(LEFT(Anrede,1))= 'F'"
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 11
        
        sSQL = "UpdATE Kunden set GESCHLECHT = 'M'"
        sSQL = sSQL & " where Ucase(LEFT(Anrede,1))= 'H'"
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 12
        
        loeschNEW "ANSCHRIF", dbWK
        
    'Kunden ende
    
    loeschNEW "umsa1", dbWK
    sSQL = "select * into umsa1 from Umsatz"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 13
    
    'Artikel
    
    lblx.Caption = "Artikel werden importiert..."
    lblx.Refresh
    
    loeschNEW "artikelK", dbWK
    
    sSQL = "Select  * into artikelK from  Artikel"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 14
    
    sSQL = "Delete from  artikelK"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 16
    
    loeschNEW "ARTIKEL", dbWK
    TransferTab dbFUT, cPfad & "KissFUT.mdb", "ARTIKEL"
    
    txtStatus.Text = 18
    
    loeschNEW "LAGER", dbWK
    TransferTab dbFUT, cPfad & "KissFUT.mdb", "LAGER"
    
    txtStatus.Text = 21
        
    sSQL = "Alter Table artikelK add  ART_refnummer Long"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 22
    
    sSQL = "Alter Table artikelK add  ART_GRPnummer Long"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 23
    
    sSQL = "insert into artikelK Select  "
    sSQL = sSQL & " ART_GRPnummer  "
    sSQL = sSQL & ", ART_refnummer"
    sSQL = sSQL & " from ARTIKEL"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = "UpdATE artikelK inner join lager on artikelk.art_refnummer = Lager.lag_refnummer"
    sSQL = sSQL & " set artikelk.bestand = lager.lag_bestand"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 27
    
    sSQL = "UpdATE artikelK set bestand = 0 where bestand < 0"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 28
    
    sSQL = "UpdATE artikelK set bestand = 0 where bestand is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 29
    
    
    loeschNEW "ART_EANS", dbWK
    TransferTab dbFUT, cPfad & "KissFUT.mdb", "ART_EANS"
    
    txtStatus.Text = 30
        
    sSQL = "Alter Table ART_EANS add  ARTNR Long"
    dbWK.Execute sSQL, dbFailOnError
        
    lblx.Caption = "DBASE ARTEAN werden importiert..."
    lblx.Refresh
    
    
    txtStatus.Text = 31
    
    loeschNEW "ARTEANDB", dbWK
    sSQL = "Select * into ARTEANDB from ARTEAN IN '" & cpfaddb & "' 'dBase IV;'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 32
    lblx.Caption = "DBASE Artikel werden importiert..."
    lblx.Refresh
    
    
    loeschNEW "ARTIKELDB", dbWK
    sSQL = "Select * into ARTIKELDB from ARTIKEL IN '" & cpfaddb & "' 'dBase IV;'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = "DBASE ARTLIEF werden importiert..."
    lblx.Refresh
    
    
    txtStatus.Text = 31
    
    loeschNEW "ARTLIEFDB", dbWK
    sSQL = "Select * into ARTLIEFDB from ARTLIEF IN '" & cpfaddb & "' 'dBase IV;'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE ART_EANS set AEA_EANCODE = val(AEA_EANCODE) "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 34
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE ART_EANS set AEA_EANCODE = '00' & AEA_EANCODE "
    sSQL = sSQL & " where len(AEA_EANCODE) = 10"
    
    txtStatus.Text = 36
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE ART_EANS set AEA_EANCODE = '0' & AEA_EANCODE "
    sSQL = sSQL & " where len(AEA_EANCODE) = 11"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 38
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    lblx.Caption = "EAN Abgleich..."
    lblx.Refresh
    
    
'    sSQL = "UpdATE ART_EANS inner join artikeldb on ART_EANS.AEA_EANCODE = artikeldb.ean"
'    sSQL = sSQL & " set ART_EANS.artnr = artikeldb.artnr "
'    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE ART_EANS inner join ARTEANDB on ART_EANS.AEA_EANCODE = ARTEANDB.ean"
    sSQL = sSQL & " set ART_EANS.artnr = ARTEANDB.artnr "
    dbWK.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 40
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE artikelK inner join ART_EANS on artikelk.art_refnummer = ART_EANS.AEA_refnummer"
    sSQL = sSQL & " set artikelK.artnr = ART_EANS.artnr "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 45
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE artikelK inner join artikeldb on artikelK.artnr = artikeldb.artnr"
    sSQL = sSQL & " set artikelK.bezeich = artikeldb.bezeich "
    sSQL = sSQL & " , artikelK.Inhalt = artikeldb.Inhalt "
    sSQL = sSQL & " , artikelK.Inhaltbez = artikeldb.Inhaltbez "
    sSQL = sSQL & " , artikelK.agn = artikeldb.agn "
    sSQL = sSQL & " , artikelK.pgn = artikeldb.pgn "
    
'    sSQL = sSQL & " , artikelK.LEKPR = artikeldb.LEKPR "
    sSQL = sSQL & " , artikelK.VKPR = artikeldb.VKPR "
    sSQL = sSQL & " , artikelK.MWST = artikeldb.MWST "
'    sSQL = sSQL & " , artikelK.LINR = artikeldb.LINR "
'    sSQL = sSQL & " , artikelK.LIBESNR = artikeldb.LIBESNR "
'    sSQL = sSQL & " , artikelK.LPZ = artikeldb.LPZ "
'    sSQL = sSQL & " , artikelK.RKZ = artikeldb.RKZ "
    sSQL = sSQL & " , artikelK.AUFDAT = artikeldb.AUFDAT "
'    sSQL = sSQL & " , artikelK.EXDAT = artikeldb.EXDAT "
'    sSQL = sSQL & " , artikelK.GRUNDPREIS = artikeldb.GRUNDPREIS "
    
    
    sSQL = sSQL & " , artikelK.gefuehrt = 'J' "
    sSQL = sSQL & " , artikelK.BONUS_OK = 'J' "
    sSQL = sSQL & " , artikelK.RABATT_OK = 'J' "
    sSQL = sSQL & " , artikelK.UMS_OK = 'J' "
    sSQL = sSQL & " , artikelK.AWM = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    

    
    sSQL = "UpdATE artikelK inner join ARTLIEFdb on artikelK.artnr = ARTLIEFdb.artnr"
    sSQL = sSQL & " set artikelK.LEKPR = ARTLIEFdb.LEKPR "
    sSQL = sSQL & " , artikelK.LINR = ARTLIEFdb.LINR "
    sSQL = sSQL & " , artikelK.LIBESNR = ARTLIEFdb.LIBESNR "
    sSQL = sSQL & " , artikelK.LPZ = ARTLIEFdb.LINIE "
    sSQL = sSQL & " , artikelK.EKPR = ARTLIEFdb.LEKPR "
    sSQL = sSQL & " , artikelK.EXDAT = ARTLIEFdb.EXDAT "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE artikelK  set artikelK.RKZ  = 'J'   "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE artikelK  set artikelK.RKZ  = 'N' where exdat is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE artikelK  set artikelK.GRUNDPREIS  = 'J'  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE artikelK  set Inhaltbez = '' where Inhaltbez is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE artikelK  set artikelK.GRUNDPREIS  = 'N' where Inhaltbez ='' "
    dbWK.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 48
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE artikelK inner join ART_EANS on artikelk.art_refnummer = ART_EANS.AEA_refnummer"
    sSQL = sSQL & " set artikelK.ean = ART_EANS.AEA_EANCODE "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 52
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE artikelK inner join ART_EANS on artikelk.art_refnummer = ART_EANS.AEA_refnummer"
    sSQL = sSQL & " set artikelK.ean2 = ART_EANS.AEA_EANCODE "
    sSQL = sSQL & " where artikelK.ean <> ART_EANS.AEA_EANCODE "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 73
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE artikelK inner join ART_EANS on artikelk.art_refnummer = ART_EANS.AEA_refnummer"
    sSQL = sSQL & " set artikelK.ean3 = ART_EANS.AEA_EANCODE "
    sSQL = sSQL & " where artikelK.ean <> ART_EANS.AEA_EANCODE and artikelK.ean2 <> ART_EANS.AEA_EANCODE"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 81
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    sSQL = "UpdATE artikelK inner join artikel on artikelK.art_refnummer = artikel.art_refnummer"
    sSQL = sSQL & " set artikelK.KVKPR1 = artikel.art_vkpreis "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 95
    
    lblx.Caption = sSQL
    lblx.Refresh
    
    loeschNEW "ART_KOPF", dbWK
    TransferTab dbFUT, cPfad & "KissFUT.mdb", "ART_KOPF"
    
    txtStatus.Text = 96
    
    'scheiß - Bezeichnungen übernehmen wir nicht mehr
    
''    sSQL = "UpdATE artikelK inner join ART_KOPF on artikelK.art_grpnummer = ART_KOPF.agr_grpnummer"
''    sSQL = sSQL & " set artikelK.bezeich = ART_KOPF.agr_Bontext "
''    dbWK.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 99
    dbWK.Close
    
    dbFUT.Close
    
    lblx.Caption = TimeValue(Now) & ": Futura Import Teil 2 ist fertig!": lblx.Refresh
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "FuturaImport2"
        Fehler.gsFehlertext = "Beim Futura Import Teil 2 ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub
Public Sub FuturaImport3(lblx As Label, cPfadFut As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    Dim dbWK        As Database
    Dim dbFUT       As Database
    Dim cPfad       As String
    Dim cOldpath    As String
    Dim cNewpath    As String
    Dim lRet        As Long
    Dim lfail       As Long
    Dim j           As Integer
    

    Screen.MousePointer = 11
    
    lblx.Caption = "Kssenjournal wird erstellt..."
    lblx.Refresh
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "WKLEER\"
    
    Set dbFUT = OpenDatabase(cPfadFut & "\FUTURA.MDB", False, False)
    Set dbWK = OpenDatabase(cPfad & "KissFUT.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    'Kassjour
    
    txtStatus.Text = 12
    
    lblx.Caption = "KAS_EX_D wird importiert..."
    lblx.Refresh
    
    loeschNEW "KAS_EX_D", dbWK
    TransferTab dbFUT, cPfad & "KissFUT.mdb", "KAS_EX_D"
    
    txtStatus.Text = 15
        
    lblx.Caption = "KASSTRNS wird importiert..."
    lblx.Refresh
    
    loeschNEW "KASSTRNS", dbWK
    TransferTab dbFUT, cPfad & "KissFUT.mdb", "KASSTRNS"
    
    txtStatus.Text = 19
        
    sSQL = "insert into KASSTRNS Select * "
    sSQL = sSQL & " from KAS_EX_D"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 21

    loeschNEW "KAS_EX_D", dbWK
    
    sSQL = "Delete from KASSTRNS where  "
    sSQL = sSQL & "  KAS_SATZART <> 14 and KAS_SATZART <> 15 and KAS_SATZART <> 23 and KAS_SATZART <> 30 "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = "Alter Table KASSTRNS add  ARTNR Long"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table KASSTRNS add  BEZEICH TEXT(35)"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE KASSTRNS inner join artikelk on artikelk.art_refnummer = KASSTRNS.kas_refnummer"
    sSQL = sSQL & " set KASSTRNS.artnr = artikelk.artnr "
    sSQL = sSQL & " , KASSTRNS.BEZEICH = artikelk.BEZEICH "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 32
    
    sSQL = "Delete from Kassjour"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update KASSTRNS Set "
    sSQL = sSQL & " KAS_BETRAG  =KAS_BETRAG * KAS_ANZAHL  "
    sSQL = sSQL & " where KAS_ANZAHL > 1 "
'    sSQL = sSQL & " and KAS_satzart <> 14 "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 36
    
    sSQL = "Update KASSTRNS Set "
    sSQL = sSQL & " KAS_BETRAG  =KAS_BETRAG * (-1)  "
    sSQL = sSQL & ",KAS_ANZAHL = KAS_ANZAHL * (-1)"
    sSQL = sSQL & " where KAS_RETOUR = 1 "
    
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 44
    
    sSQL = "Update KASSTRNS Set "
    sSQL = sSQL & " KAS_BETRAG  =KAS_BETRAG * (-1)  "
    sSQL = sSQL & ",KAS_ANZAHL = KAS_ANZAHL * (-1)"
    sSQL = sSQL & " where KAS_SATZART = 30 "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 51
    
    
    
'    sSQL = "insert into KASSJOUR Select "
'    sSQL = sSQL & " ARTNR "
'    sSQL = sSQL & ", BEZEICH "
'    sSQL = sSQL & ", KAS_BETRAG * (-1) as Preis "
'    sSQL = sSQL & ", KAS_ANZAHL * (-1) as MENGE"
'    sSQL = sSQL & ", KAS_BONNR as BELEGNR"
'    sSQL = sSQL & ", KAS_ZEIT as AZEIT"
'    sSQL = sSQL & ", KAS_DATUM as ADATE"
'    sSQL = sSQL & ", 0 as FILIALE"
'    sSQL = sSQL & ", 0 as KUNDNR"
'    sSQL = sSQL & ", 1 as KASNUM"
'    sSQL = sSQL & ", 99 as Bediener"
'    sSQL = sSQL & ", 'BA' as KK_ART"
'    sSQL = sSQL & " from KASSTRNS"
'    sSQL = sSQL & " where KAS_RETOUR = 1 "
'    lblx.Caption = sSQL: lblx.Refresh
'    dbWk.Execute sSQL, dbFailOnError
    
    sSQL = "insert into KASSJOUR Select "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", KAS_BETRAG as Preis "
    sSQL = sSQL & ", KAS_ANZAHL as MENGE"
    sSQL = sSQL & ", KAS_BONNR as BELEGNR"
    sSQL = sSQL & ", KAS_ZEIT as AZEIT"
    sSQL = sSQL & ", KAS_DATUM as ADATE"
    sSQL = sSQL & ", 0 as FILIALE"
    sSQL = sSQL & ", 0 as KUNDNR"
    sSQL = sSQL & ", 1 as KASNUM"
    sSQL = sSQL & ", 99 as Bediener"
    sSQL = sSQL & ", 'BA' as KK_ART"
    sSQL = sSQL & " from KASSTRNS"
'    sSQL = sSQL & " where KAS_RETOUR = 0 "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 56
    
    sSQL = "UpdATE KASSJOUR inner join artikelK on artikelK.artnr = KASSJOUR.artnr"
    sSQL = sSQL & " set KASSJOUR.agn = artikelK.agn "
    sSQL = sSQL & " , KASSJOUR.LINR = artikelK.LINR "
    sSQL = sSQL & " , KASSJOUR.LPZ = artikelK.LPZ "
    sSQL = sSQL & " , KASSJOUR.EAN = artikelK.EAN "
    sSQL = sSQL & " , KASSJOUR.MWST = artikelK.MWST "
    sSQL = sSQL & " , KASSJOUR.EKPR = artikelK.EKPR "
    sSQL = sSQL & " , KASSJOUR.VKPR = artikelK.VKPR "
    sSQL = sSQL & " , KASSJOUR.UMS_OK = artikelK.UMS_OK "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 58
    
    sSQL = "UpdATE KASSJOUR "
    sSQL = sSQL & " set KASSJOUR.artnr = 999999 "
    sSQL = sSQL & " , KASSJOUR.BEZEICH = 'unbekannter Artikel' "
    sSQL = sSQL & " where KASSJOUR.BEZEICH is null "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 62
    
    sSQL = "UpdATE KASSJOUR "
    sSQL = sSQL & " set KASSJOUR.artnr = 999999 "
    sSQL = sSQL & " where KASSJOUR.artnr is null "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 64
    
    sSQL = "UpdATE KASSJOUR "
    sSQL = sSQL & " set KASSJOUR.UMS_OK  = 'J' "
    sSQL = sSQL & " where KASSJOUR.UMS_OK is null "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 66
    
    If NewTableSuchenDBKombi("Umsatz", dbFUT) = True Then
    
        lblx.Caption = "Kundenumsätze werden importiert..."
        lblx.Refresh
        
        loeschNEW "Umsatz", dbWK
        TransferTab dbFUT, cPfad & "KissFUT.mdb", "Umsatz"
        
        loeschNEW "KDUMS", dbWK
        sSQL = "select * into KDUMS from Umsatz"
        lblx.Caption = sSQL: lblx.Refresh
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 67
        
        sSQL = "UpdATE KASSJOUR inner join KDUMS on KASSJOUR.BELEGNR = KDUMS.UMS_BELEG_NUMMER"
        sSQL = sSQL & " and KASSJOUR.adate = KDUMS.UMS_DATUM "
        sSQL = sSQL & " set KASSJOUR.KUNDNR = KDUMS.UMS_NUMMER "
        lblx.Caption = sSQL: lblx.Refresh
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = 69
    End If
    
    loeschNEW "Umsatz", dbWK
    sSQL = "select * into Umsatz from Umsa1"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 70
    
    loeschNEW "umsa1", dbWK
    
    sSQL = "Delete from Umsatz"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Umsatz select sum(preis) as umsg1, adate as datum"
    sSQL = sSQL & " from kassjour group by adate "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 72
    
    Dim rsrs As Recordset
    Dim rsrs1 As Recordset
    Dim dKundzahl As Long
    
    sSQL = "select * from umsatz "
    Set rsrs = dbWK.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!Datum) Then
                dKundzahl = 0
                sSQL = "select distinct(belegnr) as maxi "
                sSQL = sSQL & " from kassjour where adate = " & CLng(rsrs!Datum)
                sSQL = sSQL & " and ums_ok = 'J'"
                Set rsrs1 = dbWK.OpenRecordset(sSQL)
                If Not rsrs1.EOF Then
                    rsrs1.MoveLast
                    dKundzahl = rsrs1.RecordCount
                End If
                rsrs1.Close: Set rsrs1 = Nothing
            End If
            rsrs.Edit
            rsrs!KUNZ1 = dKundzahl
            rsrs.Update
        rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    txtStatus.Text = 74
    
    sSQL = "Delete from Artikelk where artnr is null"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Alter table Artikelk drop ART_refnummer"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78

    sSQL = "Alter table Artikelk drop ART_grpnummer"
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError

    loeschNEW "Artikel", dbWK

    sSQL = "select * into ARTIKEL from Artikelk "
    lblx.Caption = sSQL: lblx.Refresh
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 82
    
    dbWK.Close
    
    dbFUT.Close
    
    
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "WKLEER\"
    
    cOldpath = cPfad
    cOldpath = ShortPath(cOldpath)
    cOldpath = cOldpath & "KissFUT.mdb"
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    txtStatus.Text = 77
    cNewpath = cPfad
    cNewpath = ShortPath(cNewpath)
    cNewpath = cNewpath & "Kissdata.mdb"
    lRet = CopyFile(cOldpath, cNewpath, lfail)
    
    If lRet = 0 Then
        Screen.MousePointer = 0
        lblx.Caption = "Abbruch"
        lblx.Refresh
        Exit Sub
    End If
    
    lblx.Caption = "Datenbank wird optimiert..."
    lblx.Refresh
    
    txtStatus.Text = 99
    
    gdBase.Close
    Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    ReIndiziereArtikelWKL00 gdBase
'    db_Reindizieren gdBase, lblx, frmWKL151.txtStatus, frmWKL151.lbl6(28)
    
    lblx.Caption = "Artikelumsätze werden erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 20
    
    UmsartjNew lblx
    
    txtStatus.Text = 80
    Ums_artNew lblx
    
    txtStatus.Text = 100

    anzeige "Erfolg", "Fertig! Ihre Daten sind übernommen.", lblx
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "FuturaImport3"
        Fehler.gsFehlertext = "Beim Futura Import Teil 3 ist ein Fehler aufgetreten."
        
        Fehlermeldung1
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
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Strichcodes ist ein Fehler aufgetreten."
    
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

