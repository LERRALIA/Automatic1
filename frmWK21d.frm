VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK21d 
   BackColor       =   &H000000C0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   3915
   ClientLeft      =   2745
   ClientTop       =   2325
   ClientWidth     =   6135
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3915
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   5595
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox txtStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin sevCommand3.Command cmdStart 
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "Starten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   1815
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "Nein"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   3960
         TabIndex        =   3
         Top             =   3000
         Width           =   1815
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "Ja"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000C0&
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
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
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
         Left            =   2520
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl1 
         BackColor       =   &H000000C0&
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
         Left            =   3480
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H000000C0&
         Caption         =   "ab Datum / Uhrzeit   :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lbl3 
         BackColor       =   &H000000C0&
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lbl6 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmWK21d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Byes As Boolean
Dim Bfilh As Boolean
Private Sub Form_Activate()
On Error GoTo LOKAL_ERROR

    If gbQZBON = True Then

        Command1(0).Visible = False
        cmdStart.Visible = False
        Command1(1).Visible = False

        Me.Refresh
        If gbBargeldEingabe = True Then
            Command1_Click 1
        Else
            Command1_Click 0
        End If
        
'        Me.Refresh
'        Command1_Click 0
        Me.Refresh
        cmdStart_Click
        Me.Refresh
    Else


    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Activate"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
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
Private Sub cmdStart_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iFilnr As Integer
    Dim iRet As Integer
    Dim cDatum As String
    Dim iFileNr As Integer
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cmdStart.Enabled = False
    cmdStart.Visible = False
    Command1(1).Visible = False
    Label2.Visible = False
    Text1(0).Visible = False
    Text1(1).Visible = False
    
    iFilnr = CInt(gcFilNr)
    If iFilnr > 0 Then
        '** Für alle Filialen einschließlich Filiale 1       27.01.2004
           
        ZentraleWillsWissen "Kassendateierstellung für die Zentrale beginnt"
        AbFil1UNDDos
        ZentraleWillsWissen "Kassendateierstellung für die Zentrale endet"
        
        If gbErfolg = False Then
            MsgBox "Der Tagesabschluß konnte nicht erstellt werden. Bitte führen Sie ihn nochmals durch!", vbOKOnly + vbCritical, "Winkiss Fehler:"
            Unload frmWK21d

            Exit Sub
        End If
    End If
    
    
    If bAbschlussjetzt = False Then
        If IsAktionZulaessig("Kassenabschluss") Then
        
            Command1(1).Visible = False
            Command1(0).Visible = False
            cmdStart.Visible = False
            
            If LoescheTagesAbschlussMODUL7(gcKasNum, picprogress, txtStatus, lbl6, lbl3, lbl1, Label1) Then
                bAbschlussjetzt = True
                Command1(1).Visible = True
                Command1(1).Caption = "Beenden"
                Command1(0).Visible = False
                cmdStart.Visible = False
                
                If gdWechselgeld > 0 Then
                    updateafcstat "WECHSEL", gdWechselgeld, gcKasNum
                    InsertWechsel gdWechselgeld, CByte(gcKasNum)
                End If
                
                gdWechselgeld = 0
                Command1_Click 1
            End If
        Else
            Command1_Click 1
        End If
    Else
        Command1_Click 1
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "cmdStart_Click"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub AbFil1UNDDos()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim lHeute      As Long
    Dim dJetzt      As Double
    Dim iFileNr     As Integer
    Dim iRet        As Integer
    Dim ctmp        As String
    Dim iFilnr      As Integer
    Dim cPfad       As String
    Dim cPfad1      As String
    Dim cDatum      As String
    Dim Fdb         As Database
    
    
    cPfad = gcDBPfad        'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    
    
    Kill cPfad & "EXPORT\FZ.mdb"
    Set Fdb = CreateDatabase(cPfad & "EXPORT\FZ.mdb", dbLangGeneral, dbVersion40)
    Fdb.Close
    
    
    
    
    cPfad1 = cPfad & "EXPORT\FZ.mdb"
    
    picprogress.Visible = True
    txtStatus.Text = CStr(1 * 100 / 26)

    lbl6.Caption = "Prüfen der Datenbank"
    lbl6.Refresh
    lbl3.Caption = "1"
    lbl3.Refresh
    lbl1.Caption = "26"
    lbl1.Refresh
    
    Label1.Caption = "von"
    Label1.Refresh
    
    Label3.Caption = "Schritt"
    Label3.Refresh
    
    iFilnr = CInt(gcFilNr)
    
    lbl6.Caption = "letzten Abschluss aktualisieren"
    lbl6.Refresh
    lbl3.Caption = "2"
    lbl3.Refresh
    DoEvents
    
    cDatum = Text1(0).Text
    If IsDate(cDatum) = False Then
        cDatum = Format(DateValue(Now), "DD.MM.YYYY")
    End If
    
'    cUhrZeit = Text1(1).Text
    
    cSQL = "Select * from LASTSEND"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    
    lbl3.Caption = "2.1"
    rsrs!FILIALE = Val(gcFilNr)
    lbl3.Caption = "2.2"
    rsrs!Datum = CLng(DateValue(cDatum)) - 1
    lbl3.Caption = "2.3"
'    rsrs!Uhrzeit = cUhrZeit
    lbl3.Caption = "2.4"
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    txtStatus.Text = CStr(2 * 100 / 26)
    
    lbl6.Caption = "Filialnummer ermitteln"
    lbl6.Refresh
    lbl3.Caption = "4"
    lbl3.Refresh
    DoEvents
    
    gcFilNr = "-1"

    Set rsrs = gdBase.OpenRecordset("Fila", dbOpenTable)
    If Not rsrs.EOF Then
        gcFilNr = rsrs!fil
    End If
    rsrs.Close: Set rsrs = Nothing

    gbFilNr = False
    If Val(gcFilNr) > -1 Then
        gbFilNr = True
    End If
      
    lbl6.Caption = "Kundenausgangsdatei erzeugen 1/4"
    lbl6.Refresh
    lbl3.Caption = "6"
    lbl3.Refresh
    DoEvents

    loeschNEW "Kun_out", gdBase
    
    cSQL = "Select * into KUN_OUT from KUNDEN where SYNSTATUS = 'A' "
    cSQL = cSQL & " or  SYNSTATUS = 'E' "
    cSQL = cSQL & " or  SYNSTATUS = 'D' "
    gdBase.Execute cSQL, dbFailOnError
    
    lbl6.Caption = "Kundenausgangsdatei erzeugen 2/4"
    lbl6.Refresh
    
    cSQL = "Update KUNDEN set STATUS = 'N' where STATUS <> 'N'"
    gdBase.Execute cSQL, dbFailOnError
    
    lbl6.Caption = "Kundenausgangsdatei erzeugen 3/4"
    lbl6.Refresh
    
    cSQL = "Update KUNDEN set SYNSTATUS = 'N' where SYNSTATUS <> 'N'"
    gdBase.Execute cSQL, dbFailOnError
    
    lbl6.Caption = "Kundenausgangsdatei erzeugen 4/4"
    lbl6.Refresh
    
'    cSQL = "Update KUNDEN set STATUS = 'N', SYNSTATUS = 'N' "
'    gdBase.Execute cSQL, dbFailOnError

    txtStatus.Text = CStr(3 * 100 / 26)
    
    TransferTab gdBase, cPfad1, "KUN_OUT"
    
    txtStatus.Text = CStr(4 * 100 / 26)
    
    lbl6.Caption = "Bedienerausgangsdatei erzeugen"
    lbl6.Refresh
    lbl3.Caption = "7"
    lbl3.Refresh
    DoEvents

    loeschNEW "BED_out", gdBase
    
    cSQL = "Select * into BED_OUT from Bedname where SYNSTATUS = 'A' "
    cSQL = cSQL & " or  SYNSTATUS = 'E' "
    cSQL = cSQL & " or  SYNSTATUS = 'D' "
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = CStr(5 * 100 / 26)
    
    TransferTab gdBase, cPfad1, "BED_OUT"
    
    txtStatus.Text = CStr(6 * 100 / 26)
    
    lbl6.Caption = "Gutschriften - Ausgangsdatei erzeugen"
    lbl6.Refresh
    lbl3.Caption = "8"
    lbl3.Refresh
    DoEvents
      
      
      
      
    If gbKL_LIVEGUTSCHEIN Then
        'bei gutschein live nichts machen
    Else
      
        '*******GUT _OUT
    
        loeschNEW "GUT_OUT", gdBase
        
        lbl6.Caption = "Gutschriften - Ausgangsdatei füllen"
        lbl6.Refresh
        lbl3.Caption = "9"
        lbl3.Refresh
        DoEvents
        
        cSQL = "Select * into GUT_OUT from Gutsch"
        gdBase.Execute cSQL, dbFailOnError
        
        txtStatus.Text = CStr(7 * 100 / 26)
        
        TransferTab gdBase, cPfad1, "GUT_OUT"
        
        loeschNEW "GUT_OUT", gdBase
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    txtStatus.Text = CStr(8 * 100 / 26)
    
    loeschNEW "LOGY", gdBase
    CreateTable "LOGY", gdBase
    
    cSQL = "Insert into LogY  select top 10 Datname,datum, " & CByte(gcFilNr) & "  as filiale from Steuerki"
    cSQL = cSQL & " where Left(datname,1) = 'Y'  order by lfnr desc"
    gdBase.Execute cSQL, dbFailOnError
    
    TransferTab gdBase, cPfad1, "LogY"
    
    
    txtStatus.Text = CStr(9 * 100 / 26)
    
    '*******KOL_OUT
   
    cDatum = DateValue(Now)
    
    ErzeugeNeueKassenDatei lbl6, lbl3, txtStatus, picprogress

    If gbErfolg = False Then
        Exit Sub
    End If
    
    lHeute = Fix(Now)
    cDatum = Format$(lHeute, "DD.MM.YYYY")
    
    lbl6.Caption = "Aktuallisieren des letzten Abschlusses"
    lbl6.Refresh
    lbl3.Caption = "26"
    lbl3.Refresh
    DoEvents
    
    cSQL = "Select * from LASTSEND"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!FILIALE = Val(gcFilNr)
    rsrs!Datum = DateValue(Now)
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
                            
    lbl6.Caption = ""
    lbl6.Refresh
    lbl1.Caption = ""
    lbl1.Refresh
    lbl3.Caption = ""
    lbl3.Refresh
    Label1.Caption = ""
    Label1.Refresh
    Label3.Caption = ""
    Label3.Refresh
       
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "AbFil1UNDDos"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten. " & lbl3.Caption

        Fehlermeldung1
    End If
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
                
    Select Case Index
    
        Case Is = 0     'Ja
            If Byes = False Then
                Screen.MousePointer = 11
                setzedrucker gcBonDrucker
                SchubladeOeffnen
                setzedrucker gcListenDrucker
'                OeffneSchubladeExplizitWKL21d
                Screen.MousePointer = 0
                lbl6.Caption = "Haben Sie heute schon Ihr Tagesprotokoll gedruckt?"
                lbl6.Refresh
                Byes = True
            Else
                Tagesabschluss
                Byes = True
            End If
        Case Is = 1     'Nein
            If Byes = False Then
                lbl6.Caption = "Haben Sie heute schon Ihr Tagesprotokoll gedruckt?"
                lbl6.Refresh
                Byes = True
            Else
                If Command1(1).Caption <> "Beenden" Then
                    If gbBargeldEingabe = True Then
                        schreibeProtoAbschluss "Kassenabschluss wurde abgebrochen---------------"
                    End If
                End If
                Byes = False
                Unload frmWK21d
'                bAbschlussjetzt = False
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tagesabschluss()
    On Error GoTo LOKAL_ERROR
                      
    If gbFilNr Then
        If gcFilNr = 0 Then ' Solounternehmen gehen diesen Weg 'Or gcFilNr = 1     27.01.2004
        
            If gbBestAkt Then
                BESTAKTweg txtStatus
            End If

            If IsAktionZulaessig("Kassenabschluss") Then
                
                If LoescheTagesAbschlussMODUL7(gcKasNum, picprogress, txtStatus, lbl6, lbl3, lbl1, Label1) Then
                    bAbschlussjetzt = True
                    
                    
                    If gdWechselgeld > 0 Then
                        updateafcstat "WECHSEL", gdWechselgeld, gcKasNum
                        InsertWechsel gdWechselgeld, CByte(gcKasNum)
                    End If
                    
                    gdWechselgeld = 0
                    
                    
                    Command1(1).Visible = True
                    Command1(1).Caption = "Beenden"
                    Command1(0).Visible = False
                    cmdStart.Visible = False
                    Command1_Click 1
                End If
            End If
        Else 'alle Filialen

            abFil1
            Label2.Visible = True
            Command1(0).Visible = False
            Command1(1).Caption = "Beenden"
            cmdStart.Visible = True
            Text1(0).Visible = True
            Text1(1).Visible = True
        End If
    End If
 
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tagesabschluss"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub abFil1()
    On Error GoTo LOKAL_ERROR

    Dim iFileNr As Integer
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cDatum As String
'    Dim cUhrZeit As String
    Dim lWert As Long
    Dim iStufe As Integer
    
    
    lbl6.Caption = "Bereitstellen der Filialdaten für das Zentralprogramm"
    lbl6.Refresh
    
    iStufe = 0
            
    If Not NewTableSuchenDBKombi("LASTSEND", gdBase) Then
        CreateTable "LASTSEND", gdBase
    End If
    
    iStufe = 1
    
    Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
    iStufe = 2
'    Text1(1).Text = Format(TimeValue(Now), "HH:MM:SS")
    iStufe = 3
    
    cSQL = "Select * from LASTSEND"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    iStufe = 4
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Datum) Then
            iStufe = 5
            cDatum = rsrs!Datum
            iStufe = 6
            If IsDate(cDatum) = False Then
                iStufe = 7
                Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
            Else
                iStufe = 8
                Text1(0).Text = Format(cDatum, "DD.MM.YYYY")
            End If
        
            
        Else
            iStufe = 9
            Text1(0).Text = Format(DateValue(Now) - 1, "DD.MM.YYYY")
        End If
        
'        If Not IsNull(rsrs!Uhrzeit) Then
'            iStufe = 10
'            Text1(1).Text = Format(TimeValue(rsrs!Uhrzeit), "HH:MM:SS")
'        Else
'            iStufe = 11
'            Text1(1).Text = Format(TimeValue(Now), "HH:MM:SS")
'        End If
    End If
    iStufe = 12
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
'    If iStufe = 10 Then
'        Text1(1).Text = Format(TimeValue(Now), "HH:MM:SS")
'        Resume Next
'    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "abFil1"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss der Filialen ist ein Fehler aufgetreten. " & iStufe
        
        Fehlermeldung1
'    End If
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lbl6
    
    Bfilh = False
    Byes = False
    
    lbl6.Caption = "Möchten Sie die Kassenschublade öffnen?"
    lbl6.Refresh
        
    Me.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
