VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL155 
   Caption         =   "Esüdro EWWS Teil 2"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL155.frx":0000
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
      Picture         =   "frmWKL155.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   1680
      TabIndex        =   8
      Top             =   240
      Width           =   1680
   End
   Begin VB.PictureBox picprogress 
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   9315
      TabIndex        =   7
      Top             =   7440
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   10680
      TabIndex        =   4
      Top             =   480
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
      Left            =   10080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   11535
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "- aktuelle KISS Stammdaten (Stada.mdb Tabellen Artikel, Artlief,Artean, Lieferanten)"
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
      TabIndex        =   9
      Top             =   1560
      Width           =   10095
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   53
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   28
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
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
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   7920
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
      Caption         =   "Esüdro EWWS Teil 2"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmWKL155"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sdbPfad As String

    Select Case Index
        Case 0
            Unload frmWKL155
        Case 6
        
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .FileName = ""
                .DialogTitle = "Wo sind die KISS - Stammdaten?"

                .Filter = "Access - Dateien (*.mdb)|Stada.mdb"
                .ShowSave

'                sdbPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
                sdbPfad = cdlopen.FileName
            End With
            
            EWWSImport2 Label1(4), sdbPfad
    End Select

err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Esüdro EWWS Teil 2 ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub EWWSImport2(lblx As Label, cpfaddb As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String

    Screen.MousePointer = 11
    
    picprogress.Visible = True
    txtStatus.Text = 7
    
    lblx.Caption = "Lieferanten werden abgeglichen..."
    lblx.Refresh
    
    loeschNEW "LISRTDB", gdBase
    
    sSQL = "Select * into LISRTDB from LIEFERANTEN IN '" & cpfaddb & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 12
    
    sSQL = "Delete from LISRT where linr in (Select linr from Lisrtdb) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LISRT Select "
    sSQL = sSQL & "  LINR "
    sSQL = sSQL & ", LIEFBEZ "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", ort as stadt "
    sSQL = sSQL & ", tel "
    sSQL = sSQL & ", faxnr as fax "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", Notizen as Notiz "
    sSQL = sSQL & " from LISRTDB "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 14
    
''    lblx.Caption = "Linien werden abgeglichen..."
''    lblx.Refresh
''
''    loeschNEW "LINBEZDB", gdBase
''    sSQL = "Select * into LINBEZDB from LINBEZ IN '" & cpfaddb & "' 'dBase IV;'"
''    gdBase.Execute sSQL, dbFailOnError
''
''    sSQL = "Delete from LINBEZ"
''    gdBase.Execute sSQL, dbFailOnError
''
''    txtStatus.Text = 17
''
''    sSQL = "Insert into LINBEZ Select "
''    sSQL = sSQL & " LINR "
''    sSQL = sSQL & ", LINBEZEICH "
''    sSQL = sSQL & ", LPZ "
''    sSQL = sSQL & ", MARKE "
''    sSQL = sSQL & ", LPZ as SORTI "
''    sSQL = sSQL & " from LINBEZDB "
''    gdBase.Execute sSQL, dbFailOnError
''
''    sSQL = "Update LINBEZ set KUERZEL = left(ucase(Marke),5)"
''    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 19
    
    lblx.Caption = "Artikel werden abgeglichen..."
    lblx.Refresh
    
    
    sSQL = "Update Artikel set AGN = AGN + 1000 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 20
    
'    sSQL = "Update AGNDBF set AGN = AGN + 1000  "
'    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 21
    
    loeschNEW "ARTIKELDB", gdBase
    sSQL = "Select * into ARTIKELDB from ARTIKEL IN '" & cpfaddb & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    If SpalteInTabellegefundenNEW("artikeldb", "GRUNDPREIS", gdBase) = False Then
        SpalteAnfuegenNEW "artikeldb", "GRUNDPREIS", "Text(1)", gdBase
    End If
    
    txtStatus.Text = 23
    
    sSQL = "Update artikeldb set GRUNDPREIS ='J'  where GP = True "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 24
    
    sSQL = "Update artikeldb set GRUNDPREIS ='N'  where GP = False"
    gdBase.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 25
    
    loeschNEW "ARTEANDB", gdBase
    sSQL = "Select * into ARTEANDB from ARTEAN IN '" & cpfaddb & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    txtStatus.Text = 26
    
    
    lblx.Caption = "Artikelvorbereitung..."
    lblx.Refresh
    
    'bei Duplikaten ohne Bestand die Ean löschen
    
    If SpalteInTabellegefundenNEW("Artikel", "vEAN", gdBase) = False Then
        SpalteAnfuegenNEW "Artikel", "vEAN", "double", gdBase
    End If
    
    sSQL = "Update Artikel set vean = val(ean)  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    If SpalteInTabellegefundenNEW("Artikel", "vEAN2", gdBase) = False Then
        SpalteAnfuegenNEW "Artikel", "vEAN2", "double", gdBase
    End If

    sSQL = "Update Artikel set vean2 = val(ean2)  "
    gdBase.Execute sSQL, dbFailOnError
    

    If SpalteInTabellegefundenNEW("Artikel", "vEAN3", gdBase) = False Then
        SpalteAnfuegenNEW "Artikel", "vEAN3", "double", gdBase
    End If
    
    sSQL = "Update Artikel set vean3 = val(ean3)  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 27
    
    
    'Artean anpassen
    
    sSQL = " Alter table ARTEANDB add vEAN double  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEANDB set vean = val(ean)  "
    sSQL = sSQL & " where not ean is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Alter table ARTEANDB add sEAN Text(13)  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEANDB set sean = vean  "
    gdBase.Execute sSQL, dbFailOnError
    
    'Ende Artean anpassen
    
    If SpalteInTabellegefundenNEW("Artikel", "sEAN", gdBase) = False Then
        SpalteAnfuegenNEW "Artikel", "sEAN", "Text(13)", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("Artikel", "sEAN2", gdBase) = False Then
        SpalteAnfuegenNEW "Artikel", "sEAN2", "Text(13)", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("Artikel", "sEAN3", gdBase) = False Then
        SpalteAnfuegenNEW "Artikel", "sEAN3", "Text(13)", gdBase
    End If
    
    
    txtStatus.Text = 28
    
    sSQL = "Update Artikel set sean3 = vean3  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set sean2 = vean2  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set sean = vean  "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 33
    
    CheckIndex "Artikel", "sean", "", gdBase
    CheckIndex "Artikel", "sean2", "", gdBase
    CheckIndex "Artikel", "sean3", "", gdBase
    CheckIndex "ARTEANDB", "sean", "", gdBase
    
    txtStatus.Text = 34
    
    If SpalteInTabellegefundenNEW("Artikel", "EWWSARTNR", gdBase) = False Then
        SpalteAnfuegenNEW "Artikel", "EWWSARTNR", "LONG", gdBase
    End If
    
    sSQL = "Update Artikel set Artikel.EWWSARTNR  = Artikel.artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "Artikel", "EWWSARTNR", "", gdBase
    
    lblx.Caption = "Übereinstimmung wird gesucht(1)..."
    lblx.Refresh
    
    sSQL = "Update Artikel inner join ARTEANDB on Artikel.sean = ARTEANDB.sean "
    sSQL = sSQL & " set Artikel.artnr  = ARTEANDB.artnr"
'    sSQL = sSQL & " , Artikel.ean  = ARTEANDB.ean"
    sSQL = sSQL & " , Artikel.notizen  = 'gefunden'"
    sSQL = sSQL & " where ARTEANDB.sean <> '0'"
    sSQL = sSQL & " and artikel.sean <> '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    lblx.Caption = "Übereinstimmung wird gesucht(2)..."
    lblx.Refresh

    sSQL = "Update Artikel inner join ARTEANDB on Artikel.sean2 = ARTEANDB.sean "
    sSQL = sSQL & " set Artikel.artnr  = ARTEANDB.artnr"
'    sSQL = sSQL & " , Artikel.ean  = ARTEANDB.ean"
    sSQL = sSQL & " , Artikel.notizen  = 'gefunden'"
    sSQL = sSQL & " where ARTEANDB.sean <> '0'"
    sSQL = sSQL & " and artikel.sean2 <> '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 36
    
    lblx.Caption = "Übereinstimmung wird gesucht(3)..."
    lblx.Refresh

    sSQL = "Update Artikel inner join ARTEANDB on Artikel.sean3 = ARTEANDB.sean "
    sSQL = sSQL & " set Artikel.artnr  = ARTEANDB.artnr"
'    sSQL = sSQL & " , Artikel.ean  = ARTEANDB.ean"
    sSQL = sSQL & " , Artikel.notizen  = 'gefunden'"
    sSQL = sSQL & " where ARTEANDB.sean <> '0'"
    sSQL = sSQL & " and artikel.sean3 <> '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    'Esüdro Tester
    loeschNEW "ESDARTIKELDB", gdBase
    sSQL = "Select * into ESDARTIKELDB from ESDARTIKEL IN '" & cpfaddb & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "ESDARTIKELDB", "ebezeich", "", gdBase
    
    sSQL = "Delete from ESDARTIKELDB where ESDARTIKELDB.ebezeich = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "ESDARTIKELDATA", gdBase
    
    sSQL = "Select * into ESDARTIKELDATA from ESDARTIKELDB "
    sSQL = sSQL & " where ESDARTIKELDB.ebezeich like '*Test*'"
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "ESDARTIKELDATA", "ebezeich", "", gdBase
    
    sSQL = "Update Artikel set notizen = '' where notizen is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    loeschNEW "ArtikelDATA", gdBase
    
    sSQL = "Select * into ArtikelDATA from Artikel "
    sSQL = sSQL & " where Artikel.bezeich like '*Test*'"
    sSQL = sSQL & " and Artikel.notizen = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "Artikel", "notizen", "", gdBase
    CheckIndex "Artikel", "bezeich", "", gdBase
    
    CheckIndex "ArtikelDATA", "notizen", "", gdBase
    CheckIndex "ArtikelDATA", "bezeich", "", gdBase
    
    sSQL = "Update ArtikelDATA inner join ESDARTIKELDATA on Ucase(Trim(ArtikelDATA.bezeich)) = Ucase(Trim(ESDARTIKELDATA.ebezeich)) "
    sSQL = sSQL & " set ArtikelDATA.artnr  = ESDARTIKELDATA.artnr"
    sSQL = sSQL & " , ArtikelDATA.notizen  = 'gefunden2'"
    sSQL = sSQL & " where ArtikelDATA.notizen = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "ArtikelDATA", "libesnr", "", gdBase
    CheckIndex "ESDARTIKELDATA", "elibesnr", "", gdBase
    
    sSQL = "Update ArtikelDATA inner join ESDARTIKELDATA on Trim(ArtikelDATA.libesnr) = Trim(ESDARTIKELDATA.elibesnr) "
    sSQL = sSQL & " set ArtikelDATA.artnr  = ESDARTIKELDATA.artnr"
    sSQL = sSQL & " , ArtikelDATA.notizen  = 'gefunden3'"
    sSQL = sSQL & " where ArtikelDATA.notizen = '' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    

    sSQL = "Update Artikel inner join ArtikelDATA on Artikel.ewwsartnr = ArtikelDATA.ewwsartnr "
    sSQL = sSQL & " set Artikel.artnr  = ArtikelDATA.artnr"
    sSQL = sSQL & " , Artikel.notizen  = ArtikelDATA.notizen"
    gdBase.Execute sSQL, dbFailOnError
    

    'Ende Esüdro Tester
    
    
    txtStatus.Text = 37
    
    lblx.Caption = "Artikeldaten werden abgeglichen..."
    lblx.Refresh
    
    sSQL = "Drop Index sEan3 on ARTIKEL"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Drop Index sEAN on ARTIKEL"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Drop Index sEAN2 on ARTIKEL"
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("Artikel", "sEAN", gdBase) = True Then
        sSQL = " Alter table Artikel drop sEAN  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Artikel", "sEAN2", gdBase) = True Then
        sSQL = " Alter table Artikel drop sEAN2  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Artikel", "sEAN3", gdBase) = True Then
        sSQL = " Alter table Artikel drop sEAN3  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Artikel", "vEAN", gdBase) = True Then
        sSQL = " Alter table Artikel drop vEAN  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Artikel", "vEAN2", gdBase) = True Then
        sSQL = " Alter table Artikel drop vEAN2  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("Artikel", "vEAN3", gdBase) = True Then
        sSQL = " Alter table Artikel drop vEAN3  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
'    'KissAGN
'    sSQL = "Update Artikel inner join disAgn on Artikel.agn = disAgn.agn "
'    sSQL = sSQL & " set Artikel.agn  = disAgn.kissagn"
'    sSQL = sSQL & " where  artikel.artnr  < 100000 "
'    gdBase.Execute sSQL, dbFailOnError
    
    


    sSQL = "UpdATE artikel inner join artikeldb on artikel.artnr = artikeldb.artnr"
    sSQL = sSQL & " set artikel.bezeich = artikeldb.bezeich "
    sSQL = sSQL & " , artikel.Inhalt = artikeldb.Inhalt "
    sSQL = sSQL & " , artikel.Inhaltbez = artikeldb.Inhaltbez "
    sSQL = sSQL & " , artikel.agn = artikeldb.agn "
    sSQL = sSQL & " , artikel.pgn = artikeldb.pgn "
'    sSQL = sSQL & " , artikel.LEKPR = artikeldb.LEKPR "
    sSQL = sSQL & " , artikel.VKPR = artikeldb.VKPR "
    sSQL = sSQL & " , artikel.MWST = artikeldb.MWST "
'    sSQL = sSQL & " , artikel.LINR = artikeldb.LINR "
'    sSQL = sSQL & " , artikel.LIBESNR = artikeldb.LIBESNR "
'    sSQL = sSQL & " , artikel.LPZ = artikeldb.LPZ "
'    sSQL = sSQL & " , artikel.RKZ = artikeldb.RKZ "
    sSQL = sSQL & " , artikel.AUFDAT = artikeldb.AUFDAT "
'    sSQL = sSQL & " , artikel.EXDAT = artikeldb.EXDAT "
    sSQL = sSQL & " , artikel.GRUNDPREIS = artikeldb.GRUNDPREIS "
'    sSQL = sSQL & " , artikel.gefuehrt = 'J' "
'    sSQL = sSQL & " , artikel.BONUS_OK = 'J' "
'    sSQL = sSQL & " , artikel.RABATT_OK = 'J' "
    sSQL = sSQL & " , artikel.UMS_OK = 'J' "
    sSQL = sSQL & " , artikel.AWM = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 41
    
    lblx.Caption = "Verkaufsdaten werden abgeglichen..."
    lblx.Refresh
    
    Dim i As Integer
    Dim j As Integer
    
    For i = Year(DateValue(Now)) To Year(DateValue(Now)) - 3 Step -1
        For j = 12 To 1 Step -1
            sSQL = "Update Kassjour inner join Artikel on Kassjour.artnr = Artikel.EWWSartnr "
            sSQL = sSQL & " set Kassjour.agn = Artikel.agn "
            sSQL = sSQL & " , Kassjour.ARTNR = Artikel.ARTNR "
            sSQL = sSQL & " , Kassjour.Bezeich = Artikel.bezeich "
            sSQL = sSQL & " , Kassjour.ekpr = Artikel.ekpr "
            sSQL = sSQL & " , Kassjour.VKPR = Artikel.VKPR "
            sSQL = sSQL & " , Kassjour.LINR = Artikel.LINR "
            sSQL = sSQL & " , Kassjour.LPZ = Artikel.LPZ "
            sSQL = sSQL & " where year(Kassjour.adate) = " & i
            sSQL = sSQL & " and month(Kassjour.adate) = " & j
            gdBase.Execute sSQL, dbFailOnError
            
            lblx.Caption = "Verkaufsdaten werden abgeglichen..." & i & "," & j
            lblx.Refresh
        Next j
    Next i
    
        
    
    
    
    txtStatus.Text = 57
    
    lblx.Caption = "Artikelgruppen werden abgeglichen..."
    lblx.Refresh

''    sSQL = "Insert into AGNDBF Select * from AGNDBFsic"
''    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "SAP", gdBase

    lblx.Caption = "Artlief wird erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 72
    
    loeschNEW "ArtliefDB", gdBase
    sSQL = "Select * into ArtliefDB from Artlief IN '" & cpfaddb & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "Artliefdb", "artnr", "", gdBase

    txtStatus.Text = 73
    
    sSQL = "Update Artlief inner join artikel on Artlief.artnr = artikel.EWWSartnr "
    sSQL = sSQL & " set Artlief.ARTNR = Artikel.ARTNR "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 74
    
    sSQL = "Update Artlief inner join Artliefdb on Artlief.artnr = Artliefdb.artnr "
    sSQL = sSQL & " set Artlief.linr = Artliefdb.linr "
    sSQL = sSQL & " , Artlief.lekpr = Artliefdb.lekpr "
    sSQL = sSQL & " , Artlief.libesnr = Artliefdb.libesnr "
    sSQL = sSQL & " , Artlief.RKZ = Artliefdb.RKZ "
    sSQL = sSQL & " , Artlief.EXDAT = Artliefdb.EXDAT "
    sSQL = sSQL & " , Artlief.minmen = Artliefdb.mm "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 75
    
    sSQL = "Update Artikel inner join Artliefdb on Artikel.artnr = Artliefdb.artnr "
    sSQL = sSQL & " set Artikel.lpz = Artliefdb.linie "
    sSQL = sSQL & " , Artikel.LINR = Artliefdb.LINR "
    sSQL = sSQL & " , Artikel.libesnr = Artliefdb.libesnr "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    sSQL = "Update Artikel set ean = val(ean)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean2 = val(ean2)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean3 = val(ean3)"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Update Artikel set ean = '0' & ean  where len(ean) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean2 = '0' & ean2  where len(ean2) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean3 = '0' & ean3  where len(ean3) = 11 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Update Artikel set ean = '' where ean = '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean2 = '' where ean2 = '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean3 = '' where ean3 = '0'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    
    
    
    
    txtStatus.Text = 76
    
    For i = Year(DateValue(Now)) To Year(DateValue(Now)) - 3 Step -1
        For j = 12 To 1 Step -1
            
            
            
            sSQL = "Update Kassjour inner join Artikel on Kassjour.artnr = Artikel.artnr "
            sSQL = sSQL & " set Kassjour.LINR = Artikel.LINR "
            sSQL = sSQL & " , Kassjour.LPZ = Artikel.LPZ "
            
            sSQL = sSQL & " where year(Kassjour.adate) = " & i
            sSQL = sSQL & " and month(Kassjour.adate) = " & j
            gdBase.Execute sSQL, dbFailOnError
            
            lblx.Caption = "Verkaufsdaten werden abgeglichen..." & i & "," & j
            lblx.Refresh
        Next j
    Next i
    
    
    
    txtStatus.Text = 77
    
    lblx.Caption = "Zugänge werden abgeglichen..."
    lblx.Refresh
    
    sSQL = "Update ZUGANG inner join artikel on zugang.artnr = artikel.EWWSartnr "
    sSQL = sSQL & " set Zugang.bezeich = artikel.bezeich "
    sSQL = sSQL & " , Zugang.ARTNR = Artikel.ARTNR "
    sSQL = sSQL & " , Zugang.linr = artikel.linr "
    sSQL = sSQL & " , Zugang.ekpr = artikel.ekpr "
    sSQL = sSQL & " , Zugang.libesnr = artikel.libesnr "
    gdBase.Execute sSQL, dbFailOnError
    
    lblx.Caption = "Zbestand wird abgeglichen..."
    lblx.Refresh
    
    sSQL = "Update ZBESTAND inner join artikel on ZBESTAND.artnr = artikel.EWWSartnr "
    sSQL = sSQL & " set  ZBESTAND.ARTNR = Artikel.ARTNR "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78
    
    sSQL = "Drop Index EWWSartnr on ARTIKEL"
    gdBase.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("Artikel", "EWWSartnr", gdBase) = True Then
        sSQL = " Alter table Artikel drop EWWSartnr  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    sSQL = "Update Artikel set notizen = etimerk "
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "ARTEANDB", gdBase
    loeschNEW "ARTIKELDB", gdBase
    loeschNEW "ARTLIEFDB", gdBase
    loeschNEW "LISRTDB", gdBase

'    lblx.Caption = "Datenbank wird optimiert..."
'    lblx.Refresh

    txtStatus.Text = 79

''    gdBase.Close
''    Set gdBase = OpenDatabase(cpfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
'
'    ReIndiziereArtikelWKL00 gdBase

    lblx.Caption = "Artikelumsätze werden erstellt..."
    lblx.Refresh

    txtStatus.Text = 80

    UmsartjNew lblx

    txtStatus.Text = 82
    Ums_artNew lblx

    txtStatus.Text = 100

    anzeige "Erfolg", "Fertig! Ihre Daten sind übernommen.", lblx

    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "EWWSImport2"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'        Resume Next

    
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
    Fehler.gsFehlertext = "Im Programmteil Esüdro EWWS Teil 2 ist ein Fehler aufgetreten."
    
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

