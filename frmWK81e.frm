VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK81e 
   BackColor       =   &H00C0C000&
   Caption         =   "Termine - Export"
   ClientHeight    =   8910
   ClientLeft      =   1935
   ClientTop       =   2475
   ClientWidth     =   11910
   Icon            =   "frmWK81e.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   3120
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Caption         =   "alle Termine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "nur neue Termine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1665
      Width           =   3375
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   9960
      TabIndex        =   1
      Top             =   7320
      Width           =   1815
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
      Caption         =   "Export"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9960
      TabIndex        =   0
      Top             =   7920
      Width           =   1815
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Mitarbeiter auswählen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Termine - Export"
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
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmWK81e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            Export_Termine
        Case 1      'Beenden
            Unload frmWK81e
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Termine-Export ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    LeseOpenings
    fuellecboBediener
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Termine-Export ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellecboBediener()
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cZiel As String
    
    Combo1.Clear
    Combo1.AddItem "alle auswählen"
    
    cSQL = "Select * from BEDTERM order by bednu desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEDNU) Then
                cFeld = rsrs!BEDNU
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = Space$(3 - Len(cFeld)) & cFeld
            
            If Not IsNull(rsrs!bedname) Then
                cFeld = rsrs!bedname
            Else
                cFeld = ""
            End If
            cFeld = Trim(cFeld)
            cZiel = cZiel & " " & cFeld
            
            Combo1.AddItem cZiel
            
'            If Combo1.Text = "" Then
'                Combo1.Text = cZiel
''                Faerbebed Trim$(Left(Combo1.Text, 3))
'            End If
                
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Combo1.Text = "alle auswählen"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecboBediener"
    Fehler.gsFehlertext = "Im Programmteil Termine-Export ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseOpenings()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cWoTag As String
    Dim iLfdNr As Integer
    Dim cVon As String
    Dim cBis As String
    Dim cZeitblock As String
    Dim iWert As Integer
    
    Dim dUhrzeit As Double
    Dim dStartzeit As Double
    Dim dZeit As Double
    Dim lcount As Long
    
    iWert = 0
    cSQL = "Select * from OPENINGS order by WOTAG, LFDNR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iWert = iWert + 1
            If Not IsNull(rsrs!WoTag) Then
                cWoTag = rsrs!WoTag
            Else
                cWoTag = "0"
            End If
            If Not IsNull(rsrs!LFDNR) Then
                iLfdNr = rsrs!LFDNR
            Else
                iLfdNr = 0
            End If
            If Not IsNull(rsrs!Von) Then
                cVon = rsrs!Von
            Else
                cVon = ""
            End If
            If Not IsNull(rsrs!Bis) Then
                cBis = rsrs!Bis
            Else
                cBis = ""
            End If
            If Not IsNull(rsrs!Zeitblock) Then
                cZeitblock = rsrs!Zeitblock
            Else
                cZeitblock = ""
            End If
            
            If cZeitblock <> "" Then
                gcZeitBlock = cZeitblock
            End If
            
            gZeiten(iWert).WoTag = Val(cWoTag)
            gZeiten(iWert).LFDNR = iLfdNr
            gZeiten(iWert).Von = cVon
            gZeiten(iWert).Bis = cBis
            gZeiten(iWert).Zeitblock = Val(cZeitblock)
            rsrs.MoveNext
        Loop
    Else
        For iWert = 1 To 21
            gZeiten(iWert).WoTag = 0
            gZeiten(iWert).LFDNR = 0
            gZeiten(iWert).Von = ""
            gZeiten(iWert).Bis = ""
            gZeiten(iWert).Zeitblock = 0
            If cZeitblock <> "" Then
                gcZeitBlock = "15"
            End If
        Next iWert
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dUhrzeit = Val(gcZeitBlock) / 1440
    gcZeitBlock = Format$(dUhrzeit, "HH:MM")
    
    
    dStartzeit = 1
    For lcount = 1 To 21        '(eine Woche mit max. 3 Öffnungszeiten pro Tag)
        If gZeiten(lcount).Von <> "" Then
            dZeit = TimeValue(gZeiten(lcount).Von)
            If dZeit < dStartzeit Then
                dStartzeit = dZeit
            End If
        End If
    Next lcount
    
    dStartzeit = dStartzeit - dUhrzeit
    gcStartZeit = Format$(dStartzeit, "HH:MM")
    
    dStartzeit = 0
    For lcount = 1 To 21
        If gZeiten(lcount).Bis <> "" Then
            dZeit = TimeValue(gZeiten(lcount).Bis)
            If dZeit > dStartzeit Then
                dStartzeit = dZeit
            End If
        End If
    Next lcount
    
    gcEndeZeit = Format$(dStartzeit, "HH:MM")
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseOpenings"
    Fehler.gsFehlertext = "Im Programmteil Termine-Export ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Export_Termine()
    On Error GoTo LOKAL_ERROR
    
    Dim dabaStat                As Database
    Dim cPfad                   As String
    Dim cSQL                    As String
    Dim rsrs                    As DAO.Recordset
    Dim rsRs2                   As DAO.Recordset
    Dim rsRs3                   As DAO.Recordset
    Dim czeit                   As String
    Dim dZeit                   As Double
    Dim dViertelStunde          As Double
    Dim cKundenName             As String
    Dim cbednu                  As String
    Dim cBedname_for_Datei      As String
    Dim cDatname                As String
    Dim sBESCHREIBUNG           As String
    Dim sOrt                    As String
    Dim sBetreff                As String
    
    Dim cNeu                    As String
    Dim bAnd                    As Boolean
    
    bAnd = False
    If Option1(0).Value = True Then
        cNeu = "nurNeue"
    Else
        cNeu = ""
    End If
    
    If Combo1.Text = "alle auswählen" Then
        cbednu = ""
        cBedname_for_Datei = ""
        cDatname = "Termine.mdb"
    Else
        cbednu = Left(Combo1.Text, 3)
        cBedname_for_Datei = Trim(Combo1.Text)
        cBedname_for_Datei = SwapStr(cBedname_for_Datei, " ", "_")
        cDatname = "Termine_" & cBedname_for_Datei & ".mdb"
    End If
    
    If cNeu = "nurNeue" Then
        cDatname = "neue_" & cDatname
    Else
        cDatname = "alle_" & cDatname
    End If
    dViertelStunde = TimeValue(gcZeitBlock)
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "BOX\"

    Kill cPfad & "TERMINE_KISS.mdb"
    
    Screen.MousePointer = 11
    
    loeschNEW "KAL_IMPORT", gdBase
    CreateTableT2 "KAL_IMPORT", gdBase
    
    cSQL = "Select BUCHUNGSNR, DATUM, MIN(UHRZEIT) as TERMIN, KABINE, "
    cSQL = cSQL & "KUNDNR, KUERZEL, BEHANDLUNG, bednu, bedname "
    cSQL = cSQL & "from TERMINE "
    
    If cbednu <> "" Then
        If bAnd = True Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " bednu = " & cbednu & " "
        bAnd = True
    End If
    
    If cNeu = "nurNeue" Then
        If bAnd = True Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & "  neu = True  "
        bAnd = True
    End If
    
    cSQL = cSQL & " group by BUCHUNGSNR, DATUM, KABINE, KUNDNR, KUERZEL, BEHANDLUNG, bednu, bedname "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    cSQL = "Select BUCHUNGSNR, MAX(UHRZEIT) as ENDE "
    cSQL = cSQL & "from TERMINE "
    
    bAnd = False
    If cbednu <> "" Then
        If bAnd = True Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " bednu = " & cbednu & " "
        bAnd = True
    End If
    
    If cNeu = "nurNeue" Then
        If bAnd = True Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & "  neu = True  "
        bAnd = True
    End If
    
    cSQL = cSQL & " group by BUCHUNGSNR"
    Set rsRs3 = gdBase.OpenRecordset(cSQL)
    
    cSQL = "Select * from KAL_IMPORT "
    Set rsRs2 = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            sBetreff = ""
            sBESCHREIBUNG = ""
            
            rsRs2.AddNew
            rsRs2!BEGINNTAM = Format(rsrs!Datum, "DD.MM.YYYY")
            rsRs2!ENDETAM = Format(rsrs!Datum, "DD.MM.YYYY")

            rsRs2!BEGINNTUM = rsrs!TERMIN
            If Not rsRs3.EOF Then
                rsRs3.MoveFirst
                Do While Not rsRs3.EOF
                    If rsRs3!BUCHUNGSNR = rsrs!BUCHUNGSNR Then
                        czeit = rsRs3!ENDE
                        dZeit = TimeValue(czeit)
                        dZeit = dZeit + dViertelStunde
                        czeit = Format$(dZeit, "HH:MM")
                        rsRs2!ENDETUM = czeit
                    End If
                    rsRs3.MoveNext
                Loop
                rsRs3.MoveFirst

            End If
            
            If Not IsNull(rsrs!Kundnr) Then
                sBetreff = WhatIsXfromKu(rsrs!Kundnr, "Name") & ", " & WhatIsXfromKu(rsrs!Kundnr, "Vorname")
            End If
            
            If Not IsNull(rsrs!bedname) Then
                sBetreff = sBetreff & " bei: " & rsrs!bedname
            End If
             
            If Not IsNull(rsrs!BEDNU) Then
                sBESCHREIBUNG = "BedienerNr: " & rsrs!BEDNU
            End If
             
            If Not IsNull(rsrs!bedname) Then
                sBESCHREIBUNG = sBESCHREIBUNG & " Bedienername: " & rsrs!bedname & vbCrLf
            End If
             
            If Not IsNull(rsrs!Behandlung) Then
                 sBESCHREIBUNG = sBESCHREIBUNG & "Behandlung: " & rsrs!Behandlung & vbCrLf
            End If
             
            If Not IsNull(rsrs!Kundnr) Then
                sBESCHREIBUNG = sBESCHREIBUNG & "Kunde: " & rsrs!Kundnr & " " & WhatIsXfromKu(rsrs!Kundnr, "Name") & ", " & WhatIsXfromKu(rsrs!Kundnr, "Vorname") & vbCrLf

                sBESCHREIBUNG = sBESCHREIBUNG & WhatIsXfromKu(rsrs!Kundnr, "PLZ") & " " & WhatIsXfromKu(rsrs!Kundnr, "Stadt") & vbCrLf

                sBESCHREIBUNG = sBESCHREIBUNG & WhatIsXfromKu(rsrs!Kundnr, "STRASSE") & vbCrLf

                sBESCHREIBUNG = sBESCHREIBUNG & "Telefon: " & WhatIsXfromKu(rsrs!Kundnr, "TEL") & vbCrLf

                sBESCHREIBUNG = sBESCHREIBUNG & "Fax: " & WhatIsXfromKu(rsrs!Kundnr, "FAXNR") & vbCrLf

                sBESCHREIBUNG = sBESCHREIBUNG & "Mobil: " & WhatIsXfromKu(rsrs!Kundnr, "Mobiltel") & vbCrLf

                sBESCHREIBUNG = sBESCHREIBUNG & "Email: " & WhatIsXfromKu(rsrs!Kundnr, "Email") & vbCrLf
            End If
            
            
            If Not IsNull(rsrs!Kabine) Then
                sOrt = rsrs!Kabine
            End If
            
            rsRs2!BETREFF = sBetreff
            rsRs2!Beschreibung = sBESCHREIBUNG
            rsRs2!ErinnerungEinAus = -1
            rsRs2!Privat = -1
            rsRs2!Ort = sOrt
            rsRs2!ZEITSPANNEZEIGENALS = 2
            rsRs2.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
    rsrs.Close: Set rsrs = Nothing
    rsRs3.Close: Set rsRs3 = Nothing
    
'    cSQL = cSQL & " BETREFF varchar(255)"
'    cSQL = cSQL & ", BEGINNTAM varchar(255)"
'    cSQL = cSQL & ", BEGINNTUM varchar(255)"
'    cSQL = cSQL & ", ENDETAM varchar(255)"
'    cSQL = cSQL & ", ENDETUM varchar(255)"
'    cSQL = cSQL & ", GanztägigesEreignis BIT "
'    cSQL = cSQL & ", ErinnerungEinAus BIT "
'    cSQL = cSQL & ", Erinnerungam varchar(255) "
'    cSQL = cSQL & ", Erinnerungum varchar(255) "
'    cSQL = cSQL & ", Besprechungsplanung varchar(255) "
'    cSQL = cSQL & ", ErforderlicheTeilnehmer varchar(255) "
'    cSQL = cSQL & ", OptionaleTeilnehmer varchar(255) "
'    cSQL = cSQL & ", Besprechungsressourcen varchar(255) "
'    cSQL = cSQL & ", Abrechnungsinformationen varchar(255) "
'    cSQL = cSQL & ", Beschreibung ntext "
'    cSQL = cSQL & ", Kategorien varchar(255) "
'    cSQL = cSQL & ", Ort varchar(255) "
'    cSQL = cSQL & ", Priorität varchar(255) "
'    cSQL = cSQL & ", Privat Bit"
'    cSQL = cSQL & ", Reisekilometer"
'    cSQL = cSQL & ", Vertraulichkeit"
'    cSQL = cSQL & ", Zeitspannezeigenals smallint"
'    cSQL = cSQL & ") "
'
    
    Kill cPfad & cDatname
    
    Set dabaStat = CreateDatabase(cPfad & cDatname, dbLangGeneral, dbVersion40)
    TransferTab gdBase, cPfad & cDatname, "KAL_IMPORT"
    dabaStat.Close
    
    Screen.MousePointer = 0
    
    

    MsgBox "Diese Datei ist unter ('" & cPfad & "') mit dem Namen: '" & cDatname & "'", vbInformation, "Winkiss Information:"
    
    Dim iRet        As Integer
    Dim ctemp       As String
    
    If cNeu = "nurNeue" Then
        ctemp = "Es wurden neue Termine ausgegeben. Möchten Sie die Neuheitendefinition zurücksetzen?"
        iRet = MsgBox(ctemp, vbInformation + vbYesNo, "Winkiss Frage:")
        
        If iRet = vbYes Then
        
            cSQL = "Update TERMINE "
            cSQL = cSQL & " set neu = False "
    
            If cbednu <> "" Then
                cSQL = cSQL & "where bednu = " & cbednu & " "
            End If
            
            gdBase.Execute cSQL, dbFailOnError
        
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Export_Termine"
        Fehler.gsFehlertext = "Im Programmteil Termine-Export ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Function fnHoleKundenNameVoll(cKdnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cKdName As String
    Dim cKdVorname As String
    
    cSQL = "Select * from KUNDEN where KUNDNR = " & cKdnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!vorname) Then
            cKdVorname = rsrs!vorname
        Else
            cKdVorname = ""
        End If
        If Not IsNull(rsrs!name) Then
            cKdName = rsrs!name
        Else
            cKdName = ""
        End If
        cKdVorname = Trim$(cKdVorname)
        cKdName = Trim$(cKdName)
        fnHoleKundenNameVoll = cKdVorname & " " & cKdName
    Else
        fnHoleKundenNameVoll = ""
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnHoleKundenNameVoll"
    Fehler.gsFehlertext = "Im Programmteil Termine ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    loeschNEW "KAL_IMPORT", gdBase
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
