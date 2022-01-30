VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL151 
   Caption         =   "Esüdro EWWS Import"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL151.frx":0000
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
      Picture         =   "frmWKL151.frx":0442
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
      Top             =   7200
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
      TabIndex        =   11
      Top             =   960
      Width           =   11535
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
      TabIndex        =   10
      Top             =   1560
      Width           =   10095
   End
   Begin VB.Label lbl6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "- die EWWS Daten "
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
      Top             =   2160
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
      Top             =   5040
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
      Top             =   6120
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
      Caption         =   "Esüdro EWWS Import"
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
Attribute VB_Name = "frmWKL151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim sPfad As String
    Dim sdbPfad As String

    Select Case index
        Case 0
            Unload frmWKL151
        Case 6
        
'            With cdlopen
'                .CancelError = True
'                On Error GoTo err
'                .DialogTitle = "Wo sind die EWWS - Dateien?"
'
'                .Filter = "dBase - Dateien (*.dbf)|EWWS0007.dbf"
'                .ShowSave
'
'                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
'            End With
'
'            EWWSImport Label1(4), sPfad
'            EWWSImport2 Label1(4), sPfad, 1, 8
            
            
            
            
            
            
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .FileName = ""
                .DialogTitle = "Wo sind die KISS - Stammdaten?"

                .Filter = "Access - Dateien (*.mdb)|Stada.mdb"
                .ShowSave

                sdbPfad = cdlopen.FileName
            End With
            
            
            
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .FileName = ""
                .DialogTitle = "Wo sind die Sortiment-Daten?"

                .Filter = "Access - Dateien (*.mdb)|Sortiment.mdb"
                .ShowSave

                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
'                sPfad = cdlopen.FileName
            End With

            
            SortimentPlusImport2 Label1(4), sPfad, sdbPfad
                        
            
'            With cdlopen
'                .CancelError = True
'                On Error GoTo err
'                .DialogTitle = "Wo sind die Dropas - Dateien?"
'
'                .Filter = "dBase - Dateien (*.dbf)|Artikel.dbf"
'                .ShowSave
'
'                sPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
'            End With
'
'            With cdlopen
'                .CancelError = True
'                On Error GoTo err
'                .FileName = ""
'                .DialogTitle = "Wo sind die KISS - Stammdaten?"
'
'                .Filter = "Access - Dateien (*.mdb)|Stada.mdb"
'                .ShowSave
'
''                sdbPfad = Left(cdlopen.FileName, Len(cdlopen.FileName) - (Len(cdlopen.FileTitle) + 1))
'                sdbPfad = cdlopen.FileName
'            End With
'
'            DropasImport Label1(4), sPfad
'            DropasImport2 Label1(4), sPfad, sdbPfad

            
    End Select

err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Esüdro EWWS Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub KundKartPruef()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    
    
    
    
    Dim rsrs            As Recordset
    Dim cSatz           As String
    
    cSQL = " Select kundkart from kunden "

    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
    
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!KUNDKART) Then
            
                cSatz = rsrs!KUNDKART
                
                rsrs.Edit
                rsrs!KUNDKART = fnMoveKundnr2EAN8(cSatz)
                rsrs.Update
                
                
            End If
            rsrs.MoveNext
        Loop

        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KundKartPruef"
    Fehler.gsFehlertext = "Im Programmteil  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub EWWSImportT(lblx As Label, sPfad As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim dbEWWS          As Database
    Dim dbQ             As Database
    Dim lcount          As Long
    Dim sTabname        As String
    
    Set dbQ = OpenDatabase(sPfad, False, False, "dBase IV;")
    
    Kill sPfad & "\EWWSsic.MDB"
    Set dbEWWS = CreateDatabase(sPfad & "\EWWSsic.MDB", dbLangGeneral, dbVersion40)
    
    txtStatus.Text = 5
    
    For lcount = 0 To 100
        txtStatus.Text = lcount
        
        If lcount < 10 Then
            sTabname = "EWWS000" & lcount
        ElseIf lcount < 100 Then
            sTabname = "EWWS00" & lcount
        End If
        
        anzeige "normal", sTabname & " wird importiert...", lblx
        DoEvents
        sSQL = "Select * into " & sTabname & " from " & sTabname & " IN '" & sPfad & "' 'dBase IV;'"
        dbEWWS.Execute sSQL, dbFailOnError
        
    Next lcount
    
    txtStatus.Text = 0
    anzeige "normal", "Teil 1 ist fertig...", lblx
    dbEWWS.Close

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "EWWSImportT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub
Private Sub DropasImport(lblx As Label, sPfad As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim dbEWWS          As Database
    Dim dbQ             As Database
    Dim lcount          As Long
    Dim sTabname(8)    As String
    
    Set dbQ = OpenDatabase(sPfad, False, False, "dBase IV;")
    
    sSQL = "Delete from Lager where Hausnr is null"
    dbQ.Execute sSQL, dbFailOnError
    
    Kill sPfad & "\Dropas.MDB"
    Set dbEWWS = CreateDatabase(sPfad & "\Dropas.MDB", dbLangGeneral, dbVersion40)
    
    picprogress.Visible = True
    txtStatus.Text = 5
    sTabname(0) = "Artikel"
    sTabname(1) = "Kunden"
    sTabname(2) = "KASSBER"
    sTabname(3) = "HAUS"
    sTabname(4) = "LAGER"
    sTabname(5) = "LIEFERAN"
    sTabname(6) = "PERSONAL"
    sTabname(7) = "KDANALYS"
'    sTabname(8) = "EWWS0019"
'    sTabname(9) = "EWWS0021"
'    sTabname(10) = "EWWS0024"
'    sTabname(11) = "EWWS0025"
'    sTabname(12) = "EWWS0068"
    
    For lcount = 0 To 7
        txtStatus.Text = 9 * lcount
        If Datendrin(sTabname(lcount), dbQ) = True Then
            anzeige "normal", sTabname(lcount) & " wird importiert...", lblx
            DoEvents
            sSQL = "Select * into DRO_" & sTabname(lcount) & " from " & sTabname(lcount) & " IN '" & sPfad & "' 'dBase IV;'"
            dbEWWS.Execute sSQL, dbFailOnError
        End If
    Next lcount
    
'    sSQL = "Select * into Lager from Lager IN '" & sPfad & "' 'dBase IV;'"
'    sSQL = sSQL & " where hausnr = 1"
'            dbEWWS.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 0
    anzeige "normal", "Teil 1 ist fertig...", lblx
    dbEWWS.Close

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "DropasImport"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub DropasImport2(lblx As Label, sPfad As String, cpfaddb As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    Dim dbWK        As Database
    Dim dbEWWS      As Database
    Dim cOldpath    As String
    Dim cNewpath    As String
    Dim lRet        As Long
    Dim lfail       As Long
    Dim j           As Integer
    Dim cpfadEWWS   As String
    Dim cPfad       As String
    Dim rsrs        As Recordset

    Screen.MousePointer = 11
    
    lblx.Caption = "Winkiss Datenbank wird erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 5
    
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
    cNewpath = cNewpath & "KissEWWS.mdb"
    lRet = CopyFile(cOldpath, cNewpath, lfail)
    

    If lRet = 0 Then
        Screen.MousePointer = 0
        lblx.Caption = "Abbruch"
        lblx.Refresh
        Exit Sub
    End If
    
    cpfadEWWS = sPfad '"C:\Daten"
    
    Set dbEWWS = OpenDatabase(cpfadEWWS & "\DROPAS.MDB", False, False)
    Set dbWK = OpenDatabase(cPfad & "KissEWWS.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    'Kunden
    txtStatus.Text = 7
    lblx.Caption = "Kunden werden importiert..."
    lblx.Refresh
    
    loeschNEW "DRO_KUNDEN", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_KUNDEN"
    
    txtStatus.Text = 9
    
    sSQL = "Delete * from kunden "
    dbWK.Execute sSQL, dbFailOnError
        
    sSQL = "Insert into Kunden Select  "
    sSQL = sSQL & " val(Kundennr) as KUNDNR "
    sSQL = sSQL & ", name2 as vorname "
    sSQL = sSQL & ", name1 as name "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", ort as stadt "
    sSQL = sSQL & ", telefon as tel "
    sSQL = sSQL & ", '' as faxnr "
    sSQL = sSQL & ", GEB_DAT as datum1"
    sSQL = sSQL & ", Rabatt "
    sSQL = sSQL & ", angelegt "
    sSQL = sSQL & ", anrede "
    sSQL = sSQL & ", Trim(KURZBEZ) & ' ' & Trim(bemerkung)  as notizen "
    sSQL = sSQL & ", Trim(lkz) as kurztext1 "
    sSQL = sSQL & ", 'N' as status "
    sSQL = sSQL & ", Kunden_kz as awm "
    sSQL = sSQL & ", titel "
    sSQL = sSQL & ", Val(Hausnr) as filialnr "
    sSQL = sSQL & " from DRO_KUNDEN "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 10
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!NOTIZEN) Then
                rsrs.Edit
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, ",", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, "'", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, ";", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, "!", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!strasse) Then
                rsrs.Edit
                rsrs!strasse = SwapStr(rsrs!strasse, ",", "")
                rsrs!strasse = SwapStr(rsrs!strasse, "'", "")
                rsrs!strasse = SwapStr(rsrs!strasse, ";", "")
                rsrs!strasse = SwapStr(rsrs!strasse, "!", "")
                rsrs!strasse = SwapStr(rsrs!strasse, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!name) Then
                rsrs.Edit
                rsrs!name = SwapStr(rsrs!name, ",", "")
                rsrs!name = SwapStr(rsrs!name, "'", "")
                rsrs!name = SwapStr(rsrs!name, ";", "")
                rsrs!name = SwapStr(rsrs!name, "!", "")
                rsrs!name = SwapStr(rsrs!name, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!vorname) Then
                rsrs.Edit
                rsrs!vorname = SwapStr(rsrs!vorname, ",", "")
                rsrs!vorname = SwapStr(rsrs!vorname, "'", "")
                rsrs!vorname = SwapStr(rsrs!vorname, ";", "")
                rsrs!vorname = SwapStr(rsrs!vorname, "!", "")
                rsrs!vorname = SwapStr(rsrs!vorname, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM = ' ' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 11
    
    sSQL = "UpdATE Kunden set Geschlecht  = 'W'"
    sSQL = sSQL & " where Ucase(Anrede) = 'FRAU'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 12
    
    sSQL = "UpdATE Kunden set Geschlecht  = 'M'"
    sSQL = sSQL & " where Ucase(Anrede) = 'HERR'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 13
    
    sSQL = "UpdATE Kunden set Firma  = Anrede"
    sSQL = sSQL & " where (Ucase(Anrede) <> 'FRAU' and Ucase(Anrede) <> 'HERR' and Ucase(Anrede) <> 'FAMILIE' ) "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 22
    
    sSQL = "UpdATE Kunden set KUERZEL = UCASE(LEFT(NAME,5))"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  Kunden"
    sSQL = sSQL & " where UCASE(Kuerzel) = 'XXXXX' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set KURZTEXT1  = 'D'"
    sSQL = sSQL & " where Ucase(KURZTEXT1) = 'DE'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set KURZTEXT1  = 'D'"
    sSQL = sSQL & " where Ucase(KURZTEXT1) = 'DE.'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set KURZTEXT1  = 'D'"
    sSQL = sSQL & " where Ucase(KURZTEXT1) = 'DEM'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set KURZTEXT1  = 'D'"
    sSQL = sSQL & " where Ucase(KURZTEXT1) = 'DED'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set KURZTEXT1  = 'D'"
    sSQL = sSQL & " where Ucase(KURZTEXT1) = 'D99'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set DATUM1  = null"
    sSQL = sSQL & " where kundnr = 3060"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set angelegt  = null"
    sSQL = sSQL & " where kundnr = 5882"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set angelegt  = null"
    sSQL = sSQL & " where kundnr = 148"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 23
 
''    loeschNEW "DRO_KUNDEN", dbWK
    'Kunden ENDE
    
    'Filialen

    lblx.Caption = "Filialen werden importiert..."
    lblx.Refresh

    loeschNEW "DRO_HAUS", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_HAUS"

    sSQL = "Insert into FILIALEN Select  "
    sSQL = sSQL & " val(HAUSNR) as FILIALNR "
    sSQL = sSQL & ", HAUS_LANG  as FILIALNAME"
    sSQL = sSQL & " from DRO_HAUS "
    dbWK.Execute sSQL, dbFailOnError

'    loeschNEW "DRO_HAUS", dbWK
    'Filialen ENDE
    
    
    
    
    
    
    'Bediener
    txtStatus.Text = 24
    
    lblx.Caption = "Bediener werden importiert..."
    lblx.Refresh
    
    loeschNEW "DRO_PERSONAL", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_PERSONAL"
        
    sSQL = "Insert into Bedname Select  "
    sSQL = sSQL & " VERKNR as bednu "
    sSQL = sSQL & ", Name2 & ' ' & name1 as bedname"
    sSQL = sSQL & ", 'KISS' as passwort "
    sSQL = sSQL & ", 9 as bediener "
    sSQL = sSQL & " from DRO_PERSONAL "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from bedname "
    sSQL = sSQL & " where bedname is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from bedname "
    sSQL = sSQL & " where bedname = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Bedname (bednu,bedName,Passwort,bediener) values ( 98,'KISS','KISS',9)  "
    dbWK.Execute sSQL, dbFailOnError
    
'    loeschNEW "DRO_PERSONAL", dbWK
        
    'Bediener ENDE
    
    
    
    
    
    
    
    
    
    

    'Artikelgruppen ENDE
    
    
    'Umsatz
    txtStatus.Text = 26

    lblx.Caption = "Umsätze werden importiert..."
    lblx.Refresh

    loeschNEW "DRO_KASSBER", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_KASSBER"
    
    loeschNEW "ZUmsatz", dbWK

    sSQL = "Create table ZUmsatz "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Datum DateTime "
    sSQL = sSQL & ", UMSG double "
    sSQL = sSQL & ", UMSV double "
    sSQL = sSQL & ", UMSE double "
    sSQL = sSQL & ", UMSO double "
    sSQL = sSQL & ", Kunz long "
    sSQL = sSQL & ", EKPR double "
    sSQL = sSQL & ", Kred double "
    sSQL = sSQL & ", FILIALE long "
    sSQL = sSQL & " ) "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Insert into ZUmsatz Select  "
    sSQL = sSQL & " Datum "
    sSQL = sSQL & ", FELD1 as KUNZ "
    sSQL = sSQL & ", FELD2 as umsg "
    sSQL = sSQL & ", FELD2 as umsv "
    sSQL = sSQL & ", hausnr as filiale "
    sSQL = sSQL & " from DRO_KASSBER "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Delete from  ZUmsatz"
    sSQL = sSQL & " where Datum is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  ZUmsatz"
    sSQL = sSQL & " where UMSG = 0 "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE ZUmsatz set kunz = 0"
    sSQL = sSQL & " where kunz is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE ZUmsatz set ekpr = 0"
    sSQL = sSQL & " where ekpr is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE ZUmsatz set kred = 0"
    sSQL = sSQL & " where kred is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE ZUmsatz set umsg = 0"
    sSQL = sSQL & " where umsg is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE ZUmsatz set umsv = 0"
    sSQL = sSQL & " where umsv is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE ZUmsatz set umse = 0"
    sSQL = sSQL & " where umse is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE ZUmsatz set umso = 0"
    sSQL = sSQL & " where umso is null "
    dbWK.Execute sSQL, dbFailOnError

    'ENDE ZUMSATZ
    loeschNEW "DRO_KASSBER", dbWK
    
    
    
   
    
    txtStatus.Text = 27
    

    
    'ARTIKEL
    
    txtStatus.Text = 28
    
    lblx.Caption = "Artikel werden importiert..."
    lblx.Refresh
    
    loeschNEW "DRO_ARTIKEL", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_ARTIKEL"
    
    lblx.Caption = "Artikel werden verarbeitet..."
    lblx.Refresh
        
    txtStatus.Text = 31
    
    sSQL = "Insert into Artikel Select "
    
    sSQL = sSQL & " Bezeichng as BEZEICH "
    sSQL = sSQL & ", Trim(artikelnr) as ean "
'    sSQL = sSQL & ", val(artgruppe) as agn "
'    sSQL = sSQL & ", val(MINIMENGE) as minmen "
'    sSQL = sSQL & ", EANNUMMER as EAN "
    sSQL = sSQL & ", EKPREIS as LEKPR "
    sSQL = sSQL & ", DSEKPREIS as EKPR "
'    sSQL = sSQL & ", KDNVKPREIS as VKPR "
    sSQL = sSQL & ", VK1 as KVKPR1 "
'    sSQL = sSQL & ", AUFNAHME as aufdat "
    sSQL = sSQL & ", trim(LIEF_ARTNR) as libesnr "
'    sSQL = sSQL & ", 1 as lpz "
'    sSQL = sSQL & ", inhalt "
    sSQL = sSQL & ", Trim(mwst_KZ) as mwst "
'    sSQL = sSQL & ", Trim(mmauswahl) as gefuehrt "
'    sSQL = sSQL & ", Trim(mmvksperre) as Preisschu "
    sSQL = sSQL & ", liefnr as linr "
'    sSQL = sSQL & ", trim(ucase(groessnbez)) as inhaltbez "
'
    sSQL = sSQL & ", trim(Raeum_KZ) as RKZ "
    sSQL = sSQL & ", EX_DATUM as EXDAT "
'    sSQL = sSQL & ", Trim(MMPREISETI) &  Trim(MMREGALETI) as notizen "
'
    sSQL = sSQL & ", Right(RABATT_KZ,1) as RABATT_OK "
'    sSQL = sSQL & ", Right(KEINBONUS,1) as BONUS_OK "
    sSQL = sSQL & ",  groesse "
    sSQL = sSQL & " from DRO_Artikel "
    
    'nur die aktiven
'    sSQL = sSQL & " where Trim(mmauswahl) = '2' "
    
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE Artikel set LINR  = LINR + 200000 "
    sSQL = sSQL & " where LINR < 400000 "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE Artikel set MWST = 'V'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'J'"
    sSQL = sSQL & " where RABATT_OK = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'J'"
    sSQL = sSQL & " where RABATT_OK = '2' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'N'"
    sSQL = sSQL & " where RABATT_OK = '1' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'J'"
    sSQL = sSQL & " where RABATT_OK = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'J'"
    sSQL = sSQL & " where RABATT_OK = ' ' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'J'"
    sSQL = sSQL & " where RABATT_OK is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where RKZ is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where RKZ = '-' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where RKZ = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where RKZ = ' ' "
    dbWK.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 33
    
    Set rsrs = dbWK.OpenRecordset("Artikel")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEZEICH) Then
                rsrs.Edit
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, ",", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "'", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, ";", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "!", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    'Ende Artikel
    
    'mit stada Artean
    lblx.Caption = "EAN werden abgeglichen..."
    lblx.Refresh
    
    loeschNEW "ARTEANDB", dbWK
    
    sSQL = "Select * into ARTEANDB from ARTEAN IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    'bei Duplikaten ohne Bestand die Ean löschen
    
    If SpalteInTabellegefundenNEW("Artikel", "vEAN", dbWK) = False Then
        SpalteAnfuegenNEW "Artikel", "vEAN", "double", dbWK
    End If
    
    sSQL = "Update Artikel set vean = val(ean)  "
    dbWK.Execute sSQL, dbFailOnError
    
    'Artean anpassen
    
    sSQL = " Alter table ARTEANDB add vEAN double  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEANDB set vean = val(ean)  "
    sSQL = sSQL & " where not ean is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = " Alter table ARTEANDB add sEAN Text(13)  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTEANDB set sean = vean  "
    dbWK.Execute sSQL, dbFailOnError
    
    'Ende Artean anpassen

    If SpalteInTabellegefundenNEW("Artikel", "sEAN", dbWK) = False Then
        SpalteAnfuegenNEW "Artikel", "sEAN", "Text(13)", dbWK
    End If
    
    sSQL = "Update Artikel set sean = vean  "
    dbWK.Execute sSQL, dbFailOnError
    
    CheckIndex "Artikel", "sean", "", dbWK
    CheckIndex "ARTEANDB", "sean", "", dbWK
    
    sSQL = "Update Artikel inner join ARTEANDB on Artikel.sean = ARTEANDB.sean "
    sSQL = sSQL & " set Artikel.artnr  = ARTEANDB.artnr"
    sSQL = sSQL & " , Artikel.notizen  = 'gefunden'"
    sSQL = sSQL & " where ARTEANDB.sean <> '0'"
    sSQL = sSQL & " and artikel.sean <> '0'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = " Alter table Artikel add LFNR autoincrement  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel "
    sSQL = sSQL & " set Artikel.artnr  = LFNR + 600000 "
    sSQL = sSQL & " where Artikel.artnr is null "
    dbWK.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("Artikel", "LFNR", dbWK) = True Then
        sSQL = " Alter table Artikel drop LFNR  "
        dbWK.Execute sSQL, dbFailOnError
    End If
    
    
    
    If SpalteInTabellegefundenNEW("Artikel", "vEAN", dbWK) = True Then
        sSQL = " Alter table Artikel drop vEAN  "
        dbWK.Execute sSQL, dbFailOnError
    End If
    
    
    
    
    'hol mal schnell auch die Artikel
    loeschNEW "ARTIKELDB", dbWK
    sSQL = "Select * into ARTIKELDB from ARTIKEL IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("artikeldb", "GRUNDPREIS", dbWK) = False Then
        SpalteAnfuegenNEW "artikeldb", "GRUNDPREIS", "Text(1)", dbWK
    End If
    
    sSQL = "Update artikeldb set GRUNDPREIS ='J'  where GP = True "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update artikeldb set GRUNDPREIS ='N'  where GP = False"
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "UpdATE artikel inner join artikeldb on artikel.artnr = artikeldb.artnr"
    sSQL = sSQL & " set artikel.bezeich = artikeldb.bezeich "
    sSQL = sSQL & " , artikel.Inhalt = artikeldb.Inhalt "
    sSQL = sSQL & " , artikel.Inhaltbez = artikeldb.Inhaltbez "
    sSQL = sSQL & " , artikel.agn = artikeldb.agn "
    sSQL = sSQL & " , artikel.pgn = artikeldb.pgn "
    sSQL = sSQL & " , artikel.VKPR = artikeldb.VKPR "
    sSQL = sSQL & " , artikel.MWST = artikeldb.MWST "
    sSQL = sSQL & " , artikel.AUFDAT = artikeldb.AUFDAT "
    sSQL = sSQL & " , artikel.GRUNDPREIS = artikeldb.GRUNDPREIS "
    sSQL = sSQL & " , artikel.UMS_OK = 'J' "
    sSQL = sSQL & " , artikel.AWM = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    'hol mal schnell auch die Artlief
    loeschNEW "ARTLIEFDB", dbWK
    sSQL = "Select * into ARTLIEFDB from ARTLIEF IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE artikel inner join ARTLIEFDB on artikel.artnr = ARTLIEFDB.artnr"
    sSQL = sSQL & " set artikel.LEKPR = ARTLIEFDB.LEKPR "
    sSQL = sSQL & " , artikel.LINR = ARTLIEFDB.LINR "
    sSQL = sSQL & " , artikel.LIBESNR = ARTLIEFDB.LIBESNR "
    sSQL = sSQL & " , artikel.LPZ = ARTLIEFDB.LINIE "
    sSQL = sSQL & " , artikel.MINMEN = ARTLIEFDB.MM "
    sSQL = sSQL & " , artikel.RKZ = ARTLIEFDB.RKZ "
    sSQL = sSQL & " , artikel.EXDAT = ARTLIEFDB.EXDAT "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where RKZ = '-' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where RKZ = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = "Artlief wird erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 34
    ArtliefReinigenkomplett lblx, dbWK
    
    
    
    'Ende mit stada Artean
    

    'ZBESTAND
    
    txtStatus.Text = 35
    lblx.Caption = "Bestände werden importiert..."
    lblx.Refresh
    
    loeschNEW "DRO_LAGER", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_LAGER"
    
    sSQL = "Create Index BESTAND on DRO_LAGER(BESTAND) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index MIND_BEST on DRO_LAGER(MIND_BEST) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRO_LAGER set MIND_BEST =0  where MIND_BEST is null"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update DRO_LAGER set BESTAND =0  where BESTAND is null"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table ZBESTAND add  EAN TEXT(13)"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index EAN on ZBESTAND(EAN) "
    dbWK.Execute sSQL, dbFailOnError
        
    sSQL = "Insert into ZBESTAND Select  "
    sSQL = sSQL & " HAUSNR as FILIALNR "
    sSQL = sSQL & " ,val(ARTIKELNR) as EAN "
    sSQL = sSQL & " ,BESTAND "
    sSQL = sSQL & " ,MIND_BEST as MINBEST "
    sSQL = sSQL & " from DRO_LAGER where Bestand > 0 or MIND_BEST > 0"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update ZBESTAND inner join Artikel on ZBESTAND.ean = Artikel.sean "
    sSQL = sSQL & " set ZBESTAND.artnr  = Artikel.artnr"
    sSQL = sSQL & " where Artikel.sean <> '0'"
    sSQL = sSQL & " and ZBESTAND.ean <> '0'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Drop Index EAN on ZBESTAND"
    dbWK.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("ZBESTAND", "EAN", dbWK) = True Then
        sSQL = " Alter table ZBESTAND drop EAN  "
        dbWK.Execute sSQL, dbFailOnError
    End If
    
    sSQL = "Delete * from ZBESTAND where Bestand = 0 and Artnr is null"
    dbWK.Execute sSQL, dbFailOnError
    
    
    

'    loeschNEW "DRO_LAGER", dbWK
        
    'ZBESTAND ENDE
    
    txtStatus.Text = 37
    
    
    
    sSQL = "UpdATE Artikel set LINR = 0"
    sSQL = sSQL & " where LINR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set PGN = 0"
    sSQL = sSQL & " where PGN is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set AGN = 0"
    sSQL = sSQL & " where AGN is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 38
    
    lblx.Caption = "Bezeichnungen werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set UMS_OK = 'J'"
    sSQL = sSQL & " ,BONUS_OK = 'J' "
    sSQL = sSQL & " ,PREISSCHU = 'N' "
    sSQL = sSQL & " ,bestand = 0"
    sSQL = sSQL & " ,gefuehrt = 'J'"
    sSQL = sSQL & " ,AWM = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set gefuehrt = 'N'"
    sSQL = sSQL & " where trim(Ucase(groesse)) = 'SG' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set gefuehrt = 'N'"
    sSQL = sSQL & " where trim(Ucase(groesse)) = 'PSG' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set groesse = ''"
    dbWK.Execute sSQL, dbFailOnError

    
    txtStatus.Text = 40
    
    lblx.Caption = "Produktlinien werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set BEZEICH = ''"
    sSQL = sSQL & " where BEZEICH is null "
    dbWK.Execute sSQL, dbFailOnError
        
    txtStatus.Text = 41
    
    lblx.Caption = "Einkaufspreise werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set LEKPR = 0 "
    sSQL = sSQL & " where LEKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set EKPR = 0 "
    sSQL = sSQL & " where EKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 42
    
    lblx.Caption = "Kassenverkaufspreise werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set KVKPR1 = 0 "
    sSQL = sSQL & " where KVKPR1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
     sSQL = "UpdATE Artikel set KVKPR1 = KVKPR1 * 1.19 "
    sSQL = sSQL & " where KVKPR1 > 0 "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set VKPR = 0 "
    sSQL = sSQL & " where VKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 43
    
    lblx.Caption = "Bestellnummern werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set LIBESNR = ''"
    sSQL = sSQL & " where LIBESNR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set MWST = 'V'"
    sSQL = sSQL & " where MWST is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 45
    
    sSQL = "UpdATE MWSTSATZ set VOLL = 19"
    sSQL = sSQL & " , ERM = 7 WHERE FurJahr=" & Year(Date)
    dbWK.Execute sSQL, dbFailOnError
    
    'ARTIKEL ENDE
    
    'Lieferanten
    
    txtStatus.Text = 64
    
    lblx.Caption = "Lieferanten werden importiert..."
    lblx.Refresh
    
    loeschNEW "DRO_LIEFERAN", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_LIEFERAN"
        
    sSQL = "Insert into LISRT Select "
    sSQL = sSQL & " LIEFNR as LINR "
    sSQL = sSQL & ", trim(Name1) as LIEFBEZ "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", Ort as stadt "
    sSQL = sSQL & ", telefon as tel "
    sSQL = sSQL & ", telefax as fax "
    sSQL = sSQL & ", KDNR as KUNDNR "
    sSQL = sSQL & ", MIND_WERT as AWERT "
    
    
    sSQL = sSQL & ", ANSP1_Name & ' ' & ANSP1_Tel & ' ' & ANSP2_Name & ' ' & ANSP2_Tel & ' ' & ADM_NAME1 & ' ' & ADM_Name2 & ' ' & ADM_STRASS & ' ' & ADM_PLZ & ' ' & ADM_ORT & ' ' & ADM_TEL & ' ' & ADM_FAX  & ' ' & ANSP3_Name & ' ' & ANSP3_Tel as Notiz "
    sSQL = sSQL & " from DRO_LIEFERAN "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE LISRT set LINR  = LINR + 200000 "
    sSQL = sSQL & " where LINR < 400000 "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 66
    
    
    
    sSQL = "Delete from LISRT where LIEFBEZ = ''"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LISRT where LIEFBEZ is null"
    dbWK.Execute sSQL, dbFailOnError
    
     Set rsrs = dbWK.OpenRecordset("LISRT")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!LIEFBEZ) Then
                rsrs.Edit
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, ",", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, "'", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, ";", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, "!", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    'mit stada Lieferanten
    lblx.Caption = "Lieferanten werden abgeglichen..."
    lblx.Refresh
    
    loeschNEW "LISRTDB", dbWK
    
    sSQL = "Select * into LISRTDB from LIEFERANTEN IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LISRT where linr in (Select linr from Lisrtdb) "
    dbWK.Execute sSQL, dbFailOnError
    
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
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE LISRT set KUERZEL = UCASE(LEFT(LIEFBEZ,5))"
    sSQL = sSQL & " where linr in (Select linr from Artikel) "
    dbWK.Execute sSQL, dbFailOnError
    
    'Ende mit stada Lieferanten
    
    
    
'    loeschNEW "DRO_LIEFERAN", dbWK
        
    'Lieferanten ENDE
    
    
    'Kassjour
    
    lblx.Caption = "Kassjour wird importiert..."
    lblx.Refresh
    
    txtStatus.Text = 67
    
    loeschNEW "BEDNAMEDB", dbWK
    
    sSQL = "Select * into BEDNAMEDB from BEDNAME IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    loeschNEW "DRO_KDANALYS", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS.mdb", "DRO_KDANALYS"
    
    txtStatus.Text = 68
    sSQL = "Insert into Kassjour Select  "
    sSQL = sSQL & " val(artikelnr) as ean "
    sSQL = sSQL & ", menge "
    sSQL = sSQL & ", betrag as PREIS "
    sSQL = sSQL & ", Datum as adate "
    sSQL = sSQL & ", 0 as Filiale "
    sSQL = sSQL & ", 1 as KASNUM "
    sSQL = sSQL & ", 'J' as UMS_OK "
    sSQL = sSQL & ", '' as BEZEICH "
    sSQL = sSQL & ", KD_NUMMER as KUNDNR "
    sSQL = sSQL & ", left(ZEIT,2) & ':' & right(ZEIT,2) as AZEIT "
    sSQL = sSQL & ", 'V' as MWST "
    sSQL = sSQL & ", VERKNR as BEDIENER "
    sSQL = sSQL & ", BON as BELEGNR "
    sSQL = sSQL & " from DRO_KDANALYS "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 69
    
    sSQL = "Create Index ean on Kassjour(ean) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index BEDIENER on Kassjour(BEDIENER) "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    txtStatus.Text = 70
    
    sSQL = "Update Kassjour inner join Artikel on Kassjour.ean = Artikel.sean "
    sSQL = sSQL & " set Kassjour.artnr  = Artikel.artnr"
    sSQL = sSQL & " where Artikel.sean <> '0'"
    sSQL = sSQL & " and Kassjour.ean <> '0'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Update Kassjour set Kassjour.artnr  = 999999 "
    sSQL = sSQL & " where artnr is null "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
'    sSQL = "Create Index adate on Kassjour(adate) "
'    dbWK.Execute sSQL, dbFailOnError
    
    For i = 1995 To 2015
    
        For m = 1 To 12
        
        
            lblx.Caption = "Kassjour wird aktualisiert..." & i & "/" & m
            lblx.Refresh
            sSQL = "Update Kassjour inner join Artikel on Kassjour.artnr = Artikel.artnr "
            sSQL = sSQL & " set Kassjour.agn = Artikel.agn "
            sSQL = sSQL & " , Kassjour.ekpr = Artikel.ekpr "
            sSQL = sSQL & " , Kassjour.VKPR = Artikel.VKPR "
            sSQL = sSQL & " , Kassjour.ean = Artikel.ean "
            sSQL = sSQL & " , Kassjour.BEZEICH = Artikel.BEZEICH "
            sSQL = sSQL & " , Kassjour.MWST = Artikel.MWST "
            sSQL = sSQL & " where month(adate) = " & m
            sSQL = sSQL & " and year(adate) = " & i
            dbWK.Execute sSQL, dbFailOnError
            
            sSQL = "Update Kassjour inner join BEDNAMEDB on Kassjour.BEDIENER = BEDNAMEDB.bednu "
            sSQL = sSQL & " set Kassjour.filiale  = BEDNAMEDB.fil"
            sSQL = sSQL & " where month(adate) = " & m
            sSQL = sSQL & " and year(adate) = " & i
            dbWK.Execute sSQL, dbFailOnError
            
            sSQL = "Update Kassjour inner join artlief on Kassjour.artnr = artlief.artnr "
            sSQL = sSQL & " set Kassjour.linr  = artlief.linr"
            sSQL = sSQL & " where month(adate) = " & m
            sSQL = sSQL & " and year(adate) = " & i
            dbWK.Execute sSQL, dbFailOnError
            
        Next m
    
    Next i
    
'    loeschNEW "DRO_KDANALYS", dbWK
    
    lblx.Caption = "Kassjour wird aktualisiert..."
    lblx.Refresh
    txtStatus.Text = 71
    
    sSQL = "Update Kassjour  "
    sSQL = sSQL & " set Kassjour.filiale  = 20"
    sSQL = sSQL & " where Kassjour.BEDIENER = 212 "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from  Kassjour  "
    sSQL = sSQL & " where year(adate) > 2015"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour  "
    sSQL = sSQL & " set Kassjour.filiale  = 1"
    sSQL = sSQL & " where Kassjour.BEDIENER = 13 "
    dbWK.Execute sSQL, dbFailOnError
    
    loeschNEW "KUNDKASS", dbWK
    CreateTable "KUNDKASS", dbWK
    
    sSQL = "Insert into KUNDKASS "
    sSQL = sSQL & " select Filiale "
    sSQL = sSQL & " ,ADATE "
    sSQL = sSQL & " ,KUNDNR "
    sSQL = sSQL & " ,ARTNR "
    sSQL = sSQL & " ,PREIS "
    sSQL = sSQL & " ,MENGE "
    sSQL = sSQL & " ,VKPR "
    sSQL = sSQL & " ,BEDIENER as BEDNR"
    sSQL = sSQL & " from kassjour where KUNDNR > 0 "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "DROP Table Bontext "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "Alter Table Kassjour drop zbonnr "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table Kassjour drop abok "
    dbWK.Execute sSQL, dbFailOnError
        
    'Kassjour ENDE
    
    'das machen wir später
    
    sSQL = "Drop Index sEAN on ARTIKEL"
    dbWK.Execute sSQL, dbFailOnError

    If SpalteInTabellegefundenNEW("Artikel", "sEAN", dbWK) = True Then
        sSQL = " Alter table Artikel drop sEAN  "
        dbWK.Execute sSQL, dbFailOnError
    End If
    
    
    loeschNEW "DRO_KUNDEN", dbWK
    loeschNEW "DRO_HAUS", dbWK
    loeschNEW "DRO_PERSONAL", dbWK
    loeschNEW "DRO_KASSBER", dbWK
    loeschNEW "DRO_ARTIKEL", dbWK
    loeschNEW "DRO_LAGER", dbWK
    loeschNEW "DRO_LIEFERAN", dbWK
    loeschNEW "DRO_KDANALYS", dbWK
    loeschNEW "ARTIKELDB", dbWK
    loeschNEW "ARTLIEFDB", dbWK
    loeschNEW "ARTEANDB", dbWK
    loeschNEW "LISRTDB", dbWK
    
    
    
    txtStatus.Text = 89
    
    lblx.Caption = "Datenbank wird kopiert..."
    lblx.Refresh
    
    dbWK.Close
    
    dbEWWS.Close
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "WKLEER\"
    
    cOldpath = cPfad
    cOldpath = ShortPath(cOldpath)
    cOldpath = cOldpath & "KissEWWS.mdb"
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    txtStatus.Text = 90
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
    
    txtStatus.Text = 91
    
    gdBase.Close
    Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    ReIndiziereArtikelWKL00 gdBase
'    db_Reindizieren gdBase, lblx, frmWKL151.txtStatus, frmWKL151.lbl6(28)
    
    lblx.Caption = "Artikelumsätze werden erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 92
    
    UmsartjNew lblx
    
    txtStatus.Text = 93
    Ums_artNew lblx
    
    Dim sTabc As String
    sTabc = kassetabcheck(gdBase, lbl6(53), lbl6(28))
    
    If sTabc = "" Then

    Else
        MsgBox "Die Tabelle " & sTabc & " wurde nicht gefunden.", vbInformation, "Winkiss Hinweis:"
'                End
    End If
    
    lbl6(53).Caption = ""
    lbl6(28).Caption = ""

    txtStatus.Text = 100

    anzeige "Erfolg", "Fertig! Ihre Daten sind übernommen.", lblx
'    lblx.Caption = "Fertig! Ihre Daten sind übernommen."
'    lblx.Refresh
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "DROPASImport2"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
    
End Sub
Private Sub SortimentPlusImport2(lblx As Label, sPfad As String, cpfaddb As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    Dim dbWK        As Database
    Dim dbEWWS      As Database
    Dim cOldpath    As String
    Dim cNewpath    As String
    Dim lRet        As Long
    Dim lfail       As Long
    Dim j           As Integer
    Dim cpfadEWWS   As String
    Dim cPfad       As String
    Dim rsrs        As Recordset

    Screen.MousePointer = 11
    
    lblx.Caption = "Winkiss Datenbank wird erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 5
    
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
    cNewpath = cNewpath & "KissSorti.mdb"
    lRet = CopyFile(cOldpath, cNewpath, lfail)
    

    If lRet = 0 Then
        Screen.MousePointer = 0
        lblx.Caption = "Abbruch"
        lblx.Refresh
        Exit Sub
    End If
    
    cpfadEWWS = sPfad '"C:\Daten"
    
    Set dbEWWS = OpenDatabase(cpfadEWWS & "\SORTIMENT.MDB", False, False)
    Set dbWK = OpenDatabase(cPfad & "KissSorti.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    'Kunden
    txtStatus.Text = 7
    lblx.Caption = "Kunden werden importiert..."
    lblx.Refresh
    
    loeschNEW "SORT_KUNDEN", dbWK
    
    sSQL = "Select * into SORT_KUNDEN from Kunden IN '" & cpfadEWWS & "\SORTIMENT.MDB" & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    
'    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "SORT_KUNDEN"
    
    txtStatus.Text = 9
    
    sSQL = "Delete * from kunden "
    dbWK.Execute sSQL, dbFailOnError
        
    sSQL = "Insert into Kunden Select  "
    sSQL = sSQL & " val(Kundennr) as KUNDNR "
    sSQL = sSQL & ", vorname "
    sSQL = sSQL & ", name "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", ort as stadt "
    sSQL = sSQL & ", telefon as tel "
    sSQL = sSQL & ", telefax as faxnr "
    
    sSQL = sSQL & ", geschl as geschlecht "
    
    
'    sSQL = sSQL & ", GEB_DAT as datum1"
    
    sSQL = sSQL & ", kassrabatt as Rabatt "
    sSQL = sSQL & ", kundeseit as angelegt "
    sSQL = sSQL & ", '' as anrede "
    sSQL = sSQL & ", Trim(bemerkung)  as notizen "
    
    sSQL = sSQL & ", bonusums as bonus "
    
    
    sSQL = sSQL & ", land as kurztext1 "
    sSQL = sSQL & ", 'N' as status "
    sSQL = sSQL & ", keinepost as awm "
    sSQL = sSQL & ", emailadr as email "
    sSQL = sSQL & ", akadtitel as titel "
    sSQL = sSQL & ", firma "
    sSQL = sSQL & ", 0 as filialnr "
    sSQL = sSQL & " from SORT_KUNDEN "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 10
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!NOTIZEN) Then
                rsrs.Edit
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, ",", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, "'", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, ";", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, "!", "")
                rsrs!NOTIZEN = SwapStr(rsrs!NOTIZEN, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!strasse) Then
                rsrs.Edit
                rsrs!strasse = SwapStr(rsrs!strasse, ",", "")
                rsrs!strasse = SwapStr(rsrs!strasse, "'", "")
                rsrs!strasse = SwapStr(rsrs!strasse, ";", "")
                rsrs!strasse = SwapStr(rsrs!strasse, "!", "")
                rsrs!strasse = SwapStr(rsrs!strasse, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!name) Then
                rsrs.Edit
                rsrs!name = SwapStr(rsrs!name, ",", "")
                rsrs!name = SwapStr(rsrs!name, "'", "")
                rsrs!name = SwapStr(rsrs!name, ";", "")
                rsrs!name = SwapStr(rsrs!name, "!", "")
                rsrs!name = SwapStr(rsrs!name, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Set rsrs = dbWK.OpenRecordset("Kunden")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!vorname) Then
                rsrs.Edit
                rsrs!vorname = SwapStr(rsrs!vorname, ",", "")
                rsrs!vorname = SwapStr(rsrs!vorname, "'", "")
                rsrs!vorname = SwapStr(rsrs!vorname, ";", "")
                rsrs!vorname = SwapStr(rsrs!vorname, "!", "")
                rsrs!vorname = SwapStr(rsrs!vorname, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    sSQL = "UpdATE Kunden set AWM = '1'   "
    sSQL = sSQL & " where AWM = 'U' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM = ' ' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 11
    
    sSQL = "UpdATE Kunden set Geschlecht  = 'W', Anrede = 'Frau'"
    sSQL = sSQL & " where Geschlecht = '2'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 12
    
    sSQL = "UpdATE Kunden set Geschlecht  = 'M', Anrede = 'Herr'"
    sSQL = sSQL & " where Geschlecht = '1'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 13
    

    txtStatus.Text = 15
    
    sSQL = "UpdATE Kunden set KUERZEL = UCASE(LEFT(NAME,5))"
    dbWK.Execute sSQL, dbFailOnError
    


    
    
    
    'Bediener
    txtStatus.Text = 22
    
    lblx.Caption = "Bediener werden importiert..."
    lblx.Refresh
    
    loeschNEW "VERK", dbWK
    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "VERK"
    
    sSQL = "Delete * from Bedname "
    dbWK.Execute sSQL, dbFailOnError
        
    sSQL = "Insert into Bedname Select  "
    sSQL = sSQL & " val(nummer) as bednu "
    sSQL = sSQL & ", Name as bedname"
    sSQL = sSQL & ", 'KISS' as passwort "
    sSQL = sSQL & ", 9 as bediener "
    sSQL = sSQL & " from VERK "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into Bedname (bednu,bedName,Passwort,bediener) values ( 98,'KISS','KISS',9)  "
    dbWK.Execute sSQL, dbFailOnError
    
    loeschNEW "VERK", dbWK
        
    'Bediener ENDE
    
    
    'Lieferanten
    
    txtStatus.Text = 23
    
    lblx.Caption = "Lieferanten werden importiert..."
    lblx.Refresh
    
    loeschNEW "LIEF", dbWK
    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "LIEF"
    
    loeschNEW "LIEFKOND", dbWK
    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "LIEFKOND"
        
    sSQL = "Insert into LISRT Select "
    sSQL = sSQL & " val(nummer) as LINR "
    sSQL = sSQL & ", trim(Name1) as LIEFBEZ "
    sSQL = sSQL & ", left(Kurz,5) as KUERZEL "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", Ort as stadt "
    sSQL = sSQL & ", telefon as tel "
    sSQL = sSQL & ", telefax as fax "
    sSQL = sSQL & ", KDNRlief as KUNDNR "
    sSQL = sSQL & " from LIEF "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update LISRT inner join LIEFKOND on  LISRT.linr = val(LIEFKOND.liefnummer)  "
    sSQL = sSQL & " set LISRT.AWERT = LIEFKOND.mindauftr   "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 66
    
    Set rsrs = dbWK.OpenRecordset("LISRT")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!LIEFBEZ) Then
                rsrs.Edit
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, ",", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, "'", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, ";", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, "!", "")
                rsrs!LIEFBEZ = SwapStr(rsrs!LIEFBEZ, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Set rsrs = dbWK.OpenRecordset("LISRT")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Kuerzel) Then
                rsrs.Edit
                rsrs!Kuerzel = SwapStr(rsrs!Kuerzel, ",", "")
                rsrs!Kuerzel = SwapStr(rsrs!Kuerzel, "'", "")
                rsrs!Kuerzel = SwapStr(rsrs!Kuerzel, ";", "")
                rsrs!Kuerzel = SwapStr(rsrs!Kuerzel, "!", "")
                rsrs!Kuerzel = SwapStr(rsrs!Kuerzel, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    'mit stada Lieferanten
    lblx.Caption = "Lieferanten werden abgeglichen..."
    lblx.Refresh
    
    loeschNEW "LISRTDB", dbWK
    
    sSQL = "Select * into LISRTDB from LIEFERANTEN IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
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
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE LISRT set KUERZEL = UCASE(LEFT(LIEFBEZ,5))"
    sSQL = sSQL & " where kuerzel is null "
    dbWK.Execute sSQL, dbFailOnError
    
    'Ende mit stada Lieferanten
    
    loeschNEW "LIEF", dbWK
    loeschNEW "LISRTDB", dbWK
    loeschNEW "LIEFKOND", dbWK
        
    'Lieferanten ENDE
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Artikelgruppen
    txtStatus.Text = 24
    
    lblx.Caption = "Artikelgruppen werden importiert..."
    lblx.Refresh
    
    loeschNEW "Artikelgruppen", dbWK
    
    sSQL = "Select * into Artikelgruppen from Artikelgruppen IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "Delete * from AGNDBF "
    dbWK.Execute sSQL, dbFailOnError
        
    sSQL = "Insert into AGNDBF Select  "
    sSQL = sSQL & " agn "
    sSQL = sSQL & ", agnbez as agtext "
    sSQL = sSQL & ", 0 as invab "
    sSQL = sSQL & " from Artikelgruppen "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    loeschNEW "Artikelgruppen", dbWK
        
    'Bediener ENDE
    
    
    
    
    

    'Artikelgruppen ENDE
    
    
    'Umsatz
    txtStatus.Text = 26

    lblx.Caption = "Umsätze werden importiert..."
    lblx.Refresh

    loeschNEW "TAGSUMM", dbWK
    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "TAGSUMM"
    
    sSQL = "Delete * from UMSATZ "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Umsatz Select  "
    sSQL = sSQL & " Datum "
    sSQL = sSQL & ", sum(anzbon) as KUNZ1 "
    sSQL = sSQL & ", sum(Satz_1)+ sum(Satz_4) as umsg1 "
    sSQL = sSQL & ", sum(Satz_1)+ sum(Satz_4) as umsv1 "
    sSQL = sSQL & " from TAGSUMM "
    sSQL = sSQL & " where filiale = '01'"
    sSQL = sSQL & " group by Datum "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Delete from Umsatz"
    sSQL = sSQL & " where Datum is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Umsatz"
    sSQL = sSQL & " where UMSG1 = 0 "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Umsatz set kunz1 = 0"
    sSQL = sSQL & " where kunz1 is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Umsatz set ekpr1 = 0"
    sSQL = sSQL & " where ekpr1 is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Umsatz set kred1 = 0"
    sSQL = sSQL & " where kred1 is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Umsatz set umsg1 = 0"
    sSQL = sSQL & " where umsg1 is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Umsatz set umsv1 = 0"
    sSQL = sSQL & " where umsv1 is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Umsatz set umse1 = 0"
    sSQL = sSQL & " where umse1 is null "
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Umsatz set umso1 = 0"
    sSQL = sSQL & " where umso1 is null "
    dbWK.Execute sSQL, dbFailOnError

    'ENDE ZUMSATZ
    loeschNEW "TAGSUMM", dbWK
    
    
    
   
    
    
    

    
    'ARTIKEL
    
    txtStatus.Text = 28
    
    lblx.Caption = "Artikel werden importiert..."
    lblx.Refresh
    
    loeschNEW "SORT_ARTIKEL", dbWK
    
    sSQL = "Select * into SORT_ARTIKEL from Artikel IN '" & cpfadEWWS & "\SORTIMENT.MDB" & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 29
    
    lblx.Caption = "Bestände werden importiert..."
    lblx.Refresh
    
    loeschNEW "ARTFILI", dbWK
    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "ARTFILI"
    
    lblx.Caption = "Artikel werden verarbeitet..."
    lblx.Refresh
        
    txtStatus.Text = 31
    
    sSQL = "Delete * from Artikel "
    dbWK.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("Artikel", "SortArtnr", dbWK) = False Then
        SpalteAnfuegenNEW "Artikel", "SortArtnr", "long", dbWK
    End If
    
    sSQL = "Insert into Artikel Select "
    sSQL = sSQL & " Bezeichner as BEZEICH "
    sSQL = sSQL & ", 0 as artnr "
    sSQL = sSQL & ", Val(Nummer) as SortArtnr "
    sSQL = sSQL & ", Val(marke) as linr "
    sSQL = sSQL & ", val(EANNUMMER) as EAN "
    sSQL = sSQL & ",  EAN2 "
    sSQL = sSQL & ", EKPREIS as LEKPR "
    sSQL = sSQL & ", EKPREIS as EKPR "
    sSQL = sSQL & ", KDNVKPREIS as KVKPR1 "
    sSQL = sSQL & ", UVP as VKPR "
    sSQL = sSQL & ", AUFNAHME as aufdat "
    sSQL = sSQL & ", trim(Refnr) as libesnr "
    sSQL = sSQL & ", 'V' as mwst "
    sSQL = sSQL & ", 'N' as gefuehrt "
'    sSQL = sSQL & ", Trim(mmvksperre) as Preisschu "
'    sSQL = sSQL & ", liefnr as linr "
'    sSQL = sSQL & ", trim(ucase(groessnbez)) as inhaltbez "
'
'    sSQL = sSQL & ", trim(Raeum_KZ) as RKZ "
    sSQL = sSQL & ", EXDATUM as EXDAT "
    sSQL = sSQL & ", aktivier as notizen "
'
'    sSQL = sSQL & ", Right(RABATT_KZ,1) as RABATT_OK "
'    sSQL = sSQL & ", Right(KEINBONUS,1) as BONUS_OK "
'    sSQL = sSQL & ",  groesse "
    sSQL = sSQL & " from SORT_ARTIKEL "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set notizen = ''   where notizen is null "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE Artikel set gefuehrt = 'J'   where notizen <> '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set notizen = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set notizen = SortArtnr "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    
    
    
   
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    txtStatus.Text = 32
    
    lblx.Caption = "Bestände werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "Update Artikel inner join ARTFILI on  Artikel.sortartnr = val(ARTFILI.nummer)  "
    sSQL = sSQL & " set Artikel.bestand = ARTFILI.bestand   "
    sSQL = sSQL & " , Artikel.minbest = ARTFILI.meldemenge   "
    sSQL = sSQL & " where ARTFILI.filiale = '01'   "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 33
    
    lblx.Caption = "EAN Behandlung..."
    lblx.Refresh

    
    sSQL = "UpdATE Artikel set ean2 = val(ean2)   "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set EAN2 = '0' & ean2  where len(ean2) = 11 "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set EAN = '0' & ean  where len(ean) = 11 "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 34
    
    lblx.Caption = "MehrEANs werden importiert..."
    lblx.Refresh
    
    loeschNEW "MehrEAN", dbWK
    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "MehrEAN"
    
    lblx.Caption = "MehrEANs werden verarbeitet..."
    lblx.Refresh
        
    txtStatus.Text = 34
    
    sSQL = "UpdATE MehrEAN set eannummer = val(eannummer)   "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE MehrEAN set eannummer = '0' & eannummer  where len(eannummer) = 11 "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE MehrEAN set Artnummer = val(Artnummer)   "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    lblx.Caption = "EAN Erkennung startet(1.Ean)..."
    lblx.Refresh
    
    loeschNEW "EAN_Erkennung", dbWK
    
    sSQL = "Create Table EAN_Erkennung (  "
    sSQL = sSQL & " sortArtnr long "
    sSQL = sSQL & ", ean Text(13) "
    sSQL = sSQL & ", kissArtnr long "
    sSQL = sSQL & " ) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into EAN_Erkennung select  "
    sSQL = sSQL & " sortArtnr "
    sSQL = sSQL & ", ean  "
    sSQL = sSQL & ", 0 as kissArtnr "
    sSQL = sSQL & " from artikel where val(ean) > 0 "
    dbWK.Execute sSQL, dbFailOnError
    
    
    txtStatus.Text = 36
    
    lblx.Caption = "EAN Erkennung startet(2.Ean)..."
    lblx.Refresh
    
    
    
    
    sSQL = "Insert into EAN_Erkennung select  "
    sSQL = sSQL & " sortArtnr "
    sSQL = sSQL & ", ean2 as ean  "
    sSQL = sSQL & ", 0 as kissArtnr "
    sSQL = sSQL & " from artikel where val(ean2) > 0 "
'    sSQL = sSQL & " and not val(ean2) in (Select val(ean) from EAN_Erkennung )  "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 37
    
    lblx.Caption = "EAN Erkennung startet(3.Ean)..."
    lblx.Refresh
    
    
    sSQL = "Insert into EAN_Erkennung select  "
    sSQL = sSQL & " Artnummer as sortArtnr "
    sSQL = sSQL & ", eannummer as ean  "
    sSQL = sSQL & ", 0 as kissArtnr "
    sSQL = sSQL & " from MehrEAN where val(eannummer) > 0 "
'    sSQL = sSQL & " and not val(eannummer) in (Select val(ean) from EAN_Erkennung )  "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 38
    
    lblx.Caption = "EAN Erkennung ..."
    lblx.Refresh
    

    sSQL = "UpdATE EAN_Erkennung set EAN = '0' & ean  where len(ean) = 11 "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 39
    
    lblx.Caption = "Duplikate raus ..."
    lblx.Refresh
    
    
    loeschNEW "alit", dbWK
    sSQL = "select count(EAN) as count ,EAN into alit from EAN_Erkennung group by EAN having count(EAN) > 1"
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update EAN_Erkennung inner join alit on  EAN_Erkennung.ean = alit.ean  "
    sSQL = sSQL & " set EAN_Erkennung.ean = ''   "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from EAN_Erkennung where ean = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    

    sSQL = "UpdATE Artikel set RKZ = 'J'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where EXDAT is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 33
    

    
    'mit stada Artean
    lblx.Caption = "EAN werden abgeglichen..."
    lblx.Refresh
    
    loeschNEW "ARTEANDB", dbWK
    
    sSQL = "Select * into ARTEANDB from ARTEAN IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    CheckIndex "EAN_Erkennung", "ean", "", dbWK
    CheckIndex "ARTEANDB", "ean", "", dbWK
    
    sSQL = "Delete * from ARTEANDB where ean is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from ARTEANDB where ean = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from ARTEANDB where ean = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update EAN_Erkennung inner join ARTEANDB on  EAN_Erkennung.ean = ARTEANDB.ean  "
    sSQL = sSQL & " set EAN_Erkennung.kissartnr = ARTEANDB.Artnr   "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel inner join EAN_Erkennung on  Artikel.sortartnr = EAN_Erkennung.sortartnr  "
    sSQL = sSQL & " set Artikel.artnr = EAN_Erkennung.kissArtnr   "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "Artikelbezeichnung wird bereinigt..."
    lblx.Refresh
    
    sSQL = "select * from Artikel where artnr = 0 "
    Set rsrs = dbWK.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEZEICH) Then
                rsrs.Edit
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, ",", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "'", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, ";", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "!", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    'Ende Artikel
    
    
    sSQL = " Alter table Artikel add LFNR autoincrement  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel "
    sSQL = sSQL & " set Artikel.artnr  = LFNR + 600000 "
    sSQL = sSQL & " where Artikel.artnr = 0 "
    dbWK.Execute sSQL, dbFailOnError
    
    If SpalteInTabellegefundenNEW("Artikel", "LFNR", dbWK) = True Then
        sSQL = " Alter table Artikel drop LFNR  "
        dbWK.Execute sSQL, dbFailOnError
    End If
    
    'bis hier
    
    
    
    
    
    
    
    
    


    
    
    
    
    'hol mal schnell auch die Artikel
    loeschNEW "ARTIKELDB", dbWK
    sSQL = "Select * into ARTIKELDB from ARTIKEL IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    CheckIndex "ARTIKELDB", "artnr", "", dbWK
    
    If SpalteInTabellegefundenNEW("artikeldb", "GRUNDPREIS", dbWK) = False Then
        SpalteAnfuegenNEW "artikeldb", "GRUNDPREIS", "Text(1)", dbWK
    End If
    
    sSQL = "Update artikeldb set GRUNDPREIS ='J'  where GP = True "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update artikeldb set GRUNDPREIS ='N'  where GP = False"
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "UpdATE artikel inner join artikeldb on artikel.artnr = artikeldb.artnr"
    sSQL = sSQL & " set artikel.bezeich = artikeldb.bezeich "
    sSQL = sSQL & " , artikel.Inhalt = artikeldb.Inhalt "
    sSQL = sSQL & " , artikel.Inhaltbez = artikeldb.Inhaltbez "
    sSQL = sSQL & " , artikel.agn = artikeldb.agn "
    sSQL = sSQL & " , artikel.pgn = artikeldb.pgn "
    sSQL = sSQL & " , artikel.VKPR = artikeldb.VKPR "
    sSQL = sSQL & " , artikel.MWST = artikeldb.MWST "
    sSQL = sSQL & " , artikel.AUFDAT = clng(datevalue(artikeldb.AUFDAT)) "
    sSQL = sSQL & " , artikel.GRUNDPREIS = artikeldb.GRUNDPREIS "
    sSQL = sSQL & " , artikel.UMS_OK = 'J' "
    sSQL = sSQL & " , artikel.AWM = '0' "
    
    sSQL = sSQL & " and artikel.artnr < 600000"
    dbWK.Execute sSQL, dbFailOnError
    
    'hol mal schnell auch die Artlief
    
    
    lblx.Caption = "Artlief Abgleich..."
    lblx.Refresh
    
    
    
    loeschNEW "ARTLIEFDB", dbWK
    sSQL = "Select * into ARTLIEFDB from ARTLIEF IN '" & cpfaddb & "'"
    dbWK.Execute sSQL, dbFailOnError
    
    CheckIndex "ARTLIEFDB", "artnr", "", dbWK
    
    sSQL = "UpdATE ARTLIEFDB set EXDAT = clng(datevalue(now))"
    sSQL = sSQL & " where EXDAT is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE artikel inner join ARTLIEFDB on artikel.artnr = ARTLIEFDB.artnr"
    sSQL = sSQL & " set artikel.LEKPR = ARTLIEFDB.LEKPR "
    sSQL = sSQL & " , artikel.LINR = ARTLIEFDB.LINR "
    sSQL = sSQL & " , artikel.LIBESNR = ARTLIEFDB.LIBESNR "
    sSQL = sSQL & " , artikel.LPZ = ARTLIEFDB.LINIE "
    sSQL = sSQL & " , artikel.MINMEN = ARTLIEFDB.MM "
    sSQL = sSQL & " , artikel.RKZ = ARTLIEFDB.RKZ "
    sSQL = sSQL & " , artikel.EXDAT = clng(datevalue(ARTLIEFDB.EXDAT)) "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "EX Abgleich..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set EXDAT = null "
    sSQL = sSQL & " where EXDAT = clng(datevalue(now)) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'J'"
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "UpdATE Artikel set RKZ = 'N'"
    sSQL = sSQL & " where EXDAT is null "
    dbWK.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = "EAN Bereinigung..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set ean = ''"
    sSQL = sSQL & " where ean is null  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set ean2 = ''"
    sSQL = sSQL & " where ean2 is null  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set ean3 = ''"
    sSQL = sSQL & " where ean3 is null  "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE Artikel set ean = ''"
    sSQL = sSQL & " where ean = '0'  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set ean2 = ''"
    sSQL = sSQL & " where ean2 = '0'  "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set ean3 = ''"
    sSQL = sSQL & " where ean3 = '0'  "
    dbWK.Execute sSQL, dbFailOnError
    
    'Ende mit stada Artean
    
    txtStatus.Text = 34
    
    lblx.Caption = "Bestellnummern werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set LIBESNR = ''"
    sSQL = sSQL & " where LIBESNR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set LINR = 600000"
    sSQL = sSQL & " where LINR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 37
    ArtliefReinigenkomplett lblx, dbWK
    lblx.Caption = "PGN werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set PGN = 0"
    sSQL = sSQL & " where PGN is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set AGN = 0"
    sSQL = sSQL & " where AGN is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set BESTAND = 0"
    sSQL = sSQL & " where BESTAND is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set MINBEST = 0"
    sSQL = sSQL & " where MINBEST is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 38
    
    lblx.Caption = "Bezeichnungen werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set UMS_OK = 'J'"
    sSQL = sSQL & " ,BONUS_OK = 'J' "
    sSQL = sSQL & " ,rabatt_OK = 'J' "
    sSQL = sSQL & " ,PREISSCHU = 'N' "
    sSQL = sSQL & " ,gefuehrt = 'J'"
    sSQL = sSQL & " ,AWM = '0' "
    sSQL = sSQL & " ,groesse = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    lblx.Caption = "Produktlinien werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set BEZEICH = ''"
    sSQL = sSQL & " where BEZEICH is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from  Artikel where BEZEICH = ''"
    dbWK.Execute sSQL, dbFailOnError
        
    txtStatus.Text = 41
    
    lblx.Caption = "Einkaufspreise werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set LEKPR = 0 "
    sSQL = sSQL & " where LEKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set EKPR = 0 "
    sSQL = sSQL & " where EKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 42
    
    lblx.Caption = "Kassenverkaufspreise werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set KVKPR1 = 0 "
    sSQL = sSQL & " where KVKPR1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set VKPR = 0 "
    sSQL = sSQL & " where VKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 43
    
    sSQL = "UpdATE Artikel set MWST = 'V'"
    sSQL = sSQL & " where MWST is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 45
    
    sSQL = "UpdATE MWSTSATZ set VOLL = 19"
    sSQL = sSQL & " , ERM = 7 WHERE FurJahr=" & Year(Date)
    dbWK.Execute sSQL, dbFailOnError
    
    'ARTIKEL ENDE
    
    'Kassjour
    
    lblx.Caption = "Kassjour wird importiert..."
    lblx.Refresh
    
    txtStatus.Text = 67
    
    loeschNEW "VERKDAT", dbWK
    TransferTab dbEWWS, cPfad & "KissSorti.mdb", "VERKDAT"
    
    txtStatus.Text = 68
    sSQL = "Insert into Kassjour Select  "
    sSQL = sSQL & " val(nummer) as ean "
    sSQL = sSQL & ", menge "
    sSQL = sSQL & ", betrag as PREIS "
    sSQL = sSQL & ", Datum as adate "
    sSQL = sSQL & ", 5 as Filiale "
    sSQL = sSQL & ", 1 as KASNUM "
    sSQL = sSQL & ", 'J' as UMS_OK "
    sSQL = sSQL & ", '' as BEZEICH "
    sSQL = sSQL & ", val(Kunde) as KUNDNR "
    sSQL = sSQL & ", zeit as AZEIT "
    sSQL = sSQL & ", 'V' as MWST "
    sSQL = sSQL & ", val(VERK) as BEDIENER "
    sSQL = sSQL & ", val(BONNR) as BELEGNR "
    sSQL = sSQL & " from VERKDAT "
    sSQL = sSQL & " where filiale = '01'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 69
    
    sSQL = "Create Index ean on Kassjour(ean) "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 70
    
    sSQL = "Update Kassjour inner join Artikel on val(Kassjour.ean) = artikel.sortartnr "
    sSQL = sSQL & " set Kassjour.artnr  = artikel.artnr"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour set Kassjour.artnr  = 999999 "
    sSQL = sSQL & " where artnr is null "
    dbWK.Execute sSQL, dbFailOnError
    
    For i = 2006 To 2018
    
        For m = 1 To 12
        
            lblx.Caption = "Kassjour wird aktualisiert..." & i & "/" & m
            lblx.Refresh
            sSQL = "Update Kassjour inner join Artikel on Kassjour.artnr = Artikel.artnr "
            sSQL = sSQL & " set Kassjour.agn = Artikel.agn "
            sSQL = sSQL & " , Kassjour.ekpr = Artikel.ekpr "
            sSQL = sSQL & " , Kassjour.VKPR = Artikel.VKPR "
            sSQL = sSQL & " , Kassjour.linr = Artikel.Linr "
            sSQL = sSQL & " , Kassjour.lpz = Artikel.lpz "
            sSQL = sSQL & " , Kassjour.ean = Artikel.ean "
            sSQL = sSQL & " , Kassjour.BEZEICH = Artikel.BEZEICH "
            sSQL = sSQL & " , Kassjour.MWST = Artikel.MWST "
            sSQL = sSQL & " , Kassjour.KK_ART = 'BA'"
'            sSQL = sSQL & " , Kassjour.FILIALE = 0 "
            sSQL = sSQL & " , Kassjour.best1 = 0 "
            sSQL = sSQL & " where month(adate) = " & m
            sSQL = sSQL & " and year(adate) = " & i
            dbWK.Execute sSQL, dbFailOnError
            
'            sSQL = "Update Kassjour inner join BEDNAMEDB on Kassjour.BEDIENER = BEDNAMEDB.bednu "
'            sSQL = sSQL & " set Kassjour.filiale  = BEDNAMEDB.fil"
'            sSQL = sSQL & " where month(adate) = " & m
'            sSQL = sSQL & " and year(adate) = " & i
'            dbWK.Execute sSQL, dbFailOnError
'
'            sSQL = "Update Kassjour inner join artlief on Kassjour.artnr = artlief.artnr "
'            sSQL = sSQL & " set Kassjour.linr  = artlief.linr"
'            sSQL = sSQL & " where month(adate) = " & m
'            sSQL = sSQL & " and year(adate) = " & i
'            dbWK.Execute sSQL, dbFailOnError
            
        Next m
    
    Next i
    
'    loeschNEW "DRO_KDANALYS", dbWK
    
    lblx.Caption = "Kassjour wird aktualisiert..."
    lblx.Refresh
    txtStatus.Text = 71
    
    
    
    
    
'    loeschNEW "KUNDKASS", dbWK
'    CreateTable "KUNDKASS", dbWK
'
'    sSQL = "Insert into KUNDKASS "
'    sSQL = sSQL & " select Filiale "
'    sSQL = sSQL & " ,ADATE "
'    sSQL = sSQL & " ,KUNDNR "
'    sSQL = sSQL & " ,ARTNR "
'    sSQL = sSQL & " ,PREIS "
'    sSQL = sSQL & " ,MENGE "
'    sSQL = sSQL & " ,VKPR "
'    sSQL = sSQL & " ,BEDIENER as BEDNR"
'    sSQL = sSQL & " from kassjour where KUNDNR > 0 "
'    dbWK.Execute sSQL, dbFailOnError
    
    
'    sSQL = "DROP Table Bontext "
'    dbWK.Execute sSQL, dbFailOnError
'
    
'    sSQL = "Alter Table Kassjour drop zbonnr "
'    dbWK.Execute sSQL, dbFailOnError
'
'    sSQL = "Alter Table Kassjour drop abok "
'    dbWK.Execute sSQL, dbFailOnError
        
    'Kassjour ENDE
    
    'das machen wir später
    
'    sSQL = "Drop Index sEAN on ARTIKEL"
'    dbWK.Execute sSQL, dbFailOnError

    If SpalteInTabellegefundenNEW("Artikel", "SortArtnr", dbWK) = False Then
        sSQL = " Alter table Artikel drop SortArtnr  "
        dbWK.Execute sSQL, dbFailOnError
    End If
'
'    If SpalteInTabellegefundenNEW("Artikel", "sEAN", dbWK) = True Then
'        sSQL = " Alter table Artikel drop sEAN  "
'        dbWK.Execute sSQL, dbFailOnError
'    End If
    
    
    loeschNEW "DRO_KUNDEN", dbWK
    loeschNEW "DRO_HAUS", dbWK
    loeschNEW "DRO_PERSONAL", dbWK
    loeschNEW "DRO_KASSBER", dbWK
    loeschNEW "DRO_ARTIKEL", dbWK
    loeschNEW "DRO_LAGER", dbWK
    loeschNEW "DRO_LIEFERAN", dbWK
    loeschNEW "DRO_KDANALYS", dbWK
    loeschNEW "ARTIKELDB", dbWK
    loeschNEW "ARTLIEFDB", dbWK
    loeschNEW "ARTEANDB", dbWK
    loeschNEW "LISRTDB", dbWK
    
    
    
    txtStatus.Text = 89
    
    lblx.Caption = "Datenbank wird kopiert..."
    lblx.Refresh
    
    dbWK.Close
    
    dbEWWS.Close
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "WKLEER\"
    
    cOldpath = cPfad
    cOldpath = ShortPath(cOldpath)
    cOldpath = cOldpath & "KissSorti.mdb"
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    txtStatus.Text = 90
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
    
    txtStatus.Text = 91
    
    gdBase.Close
    Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    ReIndiziereArtikelWKL00 gdBase
'    db_Reindizieren gdBase, lblx, frmWKL151.txtStatus, frmWKL151.lbl6(28)
    
    lblx.Caption = "Artikelumsätze werden erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 92
    
    UmsartjNew lblx
    
    txtStatus.Text = 93
    Ums_artNew lblx
    
    Dim sTabc As String
    sTabc = kassetabcheck(gdBase, lbl6(53), lbl6(28))
    
    If sTabc = "" Then

    Else
        MsgBox "Die Tabelle " & sTabc & " wurde nicht gefunden.", vbInformation, "Winkiss Hinweis:"
'                End
    End If
    
    lbl6(53).Caption = ""
    lbl6(28).Caption = ""

    txtStatus.Text = 100

    anzeige "Erfolg", "Fertig! Ihre Daten sind übernommen.", lblx
'    lblx.Caption = "Fertig! Ihre Daten sind übernommen."
'    lblx.Refresh
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "SortimentPlusImport2"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
    
End Sub
Private Sub EWWSImport(lblx As Label, sPfad As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim dbEWWS          As Database
    Dim dbQ             As Database
    Dim lcount          As Long
    Dim sTabname(12)    As String
    
    Set dbQ = OpenDatabase(sPfad, False, False, "dBase IV;")
    
    sSQL = "Delete from EWWS0068 where not Herkunft like 'Wareneingang*'"
    dbQ.Execute sSQL, dbFailOnError
    
    Kill sPfad & "\EWWS.MDB"
    Set dbEWWS = CreateDatabase(sPfad & "\EWWS.MDB", dbLangGeneral, dbVersion40)
    
    picprogress.Visible = True
    txtStatus.Text = 5
    sTabname(0) = "EWWS0002"
    sTabname(1) = "EWWS0005"
    sTabname(2) = "EWWS0007"
    sTabname(3) = "EWWS0008"
    sTabname(4) = "EWWS0013"
    sTabname(5) = "EWWS0014"
    sTabname(6) = "EWWS0017"
    sTabname(7) = "EWWS0018"
    sTabname(8) = "EWWS0019"
    sTabname(9) = "EWWS0021"
    sTabname(10) = "EWWS0024"
    sTabname(11) = "EWWS0025"
    sTabname(12) = "EWWS0068"
    
    For lcount = 0 To 12
        txtStatus.Text = 9 * lcount
        If Datendrin(sTabname(lcount), dbQ) = True Then
            anzeige "normal", sTabname(lcount) & " wird importiert...", lblx
            DoEvents
            sSQL = "Select * into " & sTabname(lcount) & " from " & sTabname(lcount) & " IN '" & sPfad & "' 'dBase IV;'"
            dbEWWS.Execute sSQL, dbFailOnError
        End If
    Next lcount
    
    txtStatus.Text = 0
    anzeige "normal", "Teil 1 ist fertig...", lblx
    dbEWWS.Close

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "EWWSImport"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub EWWSImport2(lblx As Label, sPfad As String, iFil As Integer, iBisFil As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    Dim dbWK        As Database
    Dim dbEWWS      As Database
    Dim cOldpath    As String
    Dim cNewpath    As String
    Dim lRet        As Long
    Dim lfail       As Long
    Dim j           As Integer
    Dim cpfadEWWS   As String
    Dim cPfad       As String
    Dim rsrs        As Recordset

    Screen.MousePointer = 11
    
    lblx.Caption = "Winkiss Datenbank wird erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 5
    
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
    cNewpath = cNewpath & "KissEWWS_" & iFil & ".mdb"
    lRet = CopyFile(cOldpath, cNewpath, lfail)
    

    If lRet = 0 Then
        Screen.MousePointer = 0
        lblx.Caption = "Abbruch"
        lblx.Refresh
        Exit Sub
    End If
    
    cpfadEWWS = sPfad '"C:\Daten"
    
    Set dbEWWS = OpenDatabase(cpfadEWWS & "\EWWS.MDB", False, False)
    Set dbWK = OpenDatabase(cPfad & "KissEWWS_" & iFil & ".mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    'Kunden
    txtStatus.Text = 7
    lblx.Caption = "Kunden werden importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0018", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0018"
    
    txtStatus.Text = 9
    
    sSQL = "Delete * from kunden "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update EWWS0018 set stammfili = '0' where stammfili is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update EWWS0018 set stammfili = '0' where stammfili = '' "
    dbWK.Execute sSQL, dbFailOnError
        
    sSQL = "Insert into Kunden Select  "
    sSQL = sSQL & " val(Kundennr) as KUNDNR "
    sSQL = sSQL & ", vorName "
    sSQL = sSQL & ", name "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", ort as stadt "
    sSQL = sSQL & ", telefon as tel "
    
    sSQL = sSQL & ", telefax as faxnr "
    sSQL = sSQL & ", MOBILTEL as MOBILTEL "
    sSQL = sSQL & ", EMAILADR as EMAIL "
    sSQL = sSQL & ", BONUSUMS as BONUS "


    sSQL = sSQL & ", Rabatt "
    sSQL = sSQL & ", KUNDESEIT as angelegt "
    sSQL = sSQL & ", KUNDENART as PREISKZ "
    sSQL = sSQL & ", anrede "
    sSQL = sSQL & ", bemerkung as notizen "
    sSQL = sSQL & ", Trim(gebtag) & '.' & Trim(gebmon) & '.' & Trim(gebjahr) as kurztext1 "
    sSQL = sSQL & ", geschl as Geschlecht "
    sSQL = sSQL & ", 'N' as status "
    sSQL = sSQL & ", '0' as awm "
    sSQL = sSQL & ", '' as titel "

    sSQL = sSQL & ", Bonusmm as filialnr "
    'bei allen anderen
'    sSQL = sSQL & ", Val(Stammfili) as filialnr "
    sSQL = sSQL & " from EWWS0018 "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 10
    
    
    
    sSQL = "UpdATE Kunden inner join EWWS0018 on Kunden.kundnr = Val(EWWS0018.Kundennr)  "
    sSQL = sSQL & " set Kunden.AWM = EWWS0018.mmfrei1"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM = ' ' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set AWM = '0'   "
    sSQL = sSQL & " where AWM is null "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE Kunden set AWM = '3'   "
    sSQL = sSQL & " where val(AWM) > 0 "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    sSQL = "UpdATE Kunden set PREISKZ = 0   "
    sSQL = sSQL & " where PREISKZ = 1 "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set PREISKZ = 0   "
    sSQL = sSQL & " where PREISKZ is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set kurztext1 = ''  "
    sSQL = sSQL & " where kurztext1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set kurztext1 = ''  "
    sSQL = sSQL & " where left(kurztext1,2) = '..'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set kurztext1 = ''  "
    sSQL = sSQL & " where left(kurztext1,1) = '.'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set kurztext1 = trim(kurztext1) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Kunden set kurztext1 = ''  "
    sSQL = sSQL & " where len(kurztext1) < 10 "
    dbWK.Execute sSQL, dbFailOnError

    txtStatus.Text = 11

    loeschNEW "KUTEST", dbWK

    sSQL = "select kundnr, kurztext1 into KUTEST from Kunden "
    sSQL = sSQL & " where len(kurztext1) = 10 "
    dbWK.Execute sSQL, dbFailOnError

'diesen Befehl in Access ausführen, dann klappt es auch mit den Geburtsdaten
'update kunden inner join kutest on kunden.kundnr = kutest.kundnr set kunden.datum1 = kutest.kurztext1
    

'    sSQL = "UpdATE Kunden inner join KUTEST ON KUNDEN.KUNDNR = KUTEST.KUNDNR   "
'    sSQL = sSQL & " set KUNDEN.datum1 = KUTEST.kurztext1 "
'    dbWK.Execute sSQL, dbFailOnError
'

'    sSQL = "UpdATE Kunden set Datum1 = kurztext1 "
'    sSQL = sSQL & " where len(kurztext1) = 10 "
'    sSQL = sSQL & " and not kurztext1 is null "
'    dbWK.Execute sSQL, dbFailOnError

    
    
    
''    'manchmal auch Länge 8
''
''    sSQL = "UpdATE Kunden set notizen = notizen & kurztext1  "
''    sSQL = sSQL & " where len(kurztext1) < 8 "
''    dbWK.Execute sSQL, dbFailOnError
''
''    sSQL = "UpdATE Kunden set kurztext1 = ''  "
''    sSQL = sSQL & " where len(kurztext1) < 8 "
''    dbWK.Execute sSQL, dbFailOnError
''
''    txtStatus.Text = 11
''
''
''    sSQL = "UpdATE Kunden set Datum1 = kurztext1 "
''    sSQL = sSQL & " where len(kurztext1) = 8 "
''    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 12
    
    sSQL = "UpdATE Kunden set kurztext1 = ''  "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 13
    
    sSQL = "UpdATE Kunden set GESCHLECHT = 'W'"
    sSQL = sSQL & " where GESCHLECHT = '2'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 14
    
    sSQL = "UpdATE Kunden set Anrede = 'Frau'"
    sSQL = sSQL & " where GESCHLECHT = 'W'"
    sSQL = sSQL & " and Ucase(Anrede) <> 'FAMILIE'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 15
    
    sSQL = "UpdATE Kunden set Anrede = 'Frau'"
    sSQL = sSQL & " where GESCHLECHT = 'W'"
    sSQL = sSQL & " and Anrede = ''"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 16
    
    sSQL = "UpdATE Kunden set Anrede = 'Frau'"
    sSQL = sSQL & " where GESCHLECHT = 'W'"
    sSQL = sSQL & " and Anrede is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 17
    
    sSQL = "UpdATE Kunden set GESCHLECHT = 'M'"
    sSQL = sSQL & " where GESCHLECHT = '1'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 18
    
    sSQL = "UpdATE Kunden set Anrede = 'Herr'"
    sSQL = sSQL & " where GESCHLECHT = 'M'"
    sSQL = sSQL & " and Ucase(Anrede) <> 'FAMILIE'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 19
    
    sSQL = "UpdATE Kunden set Anrede = 'Herr'"
    sSQL = sSQL & " where GESCHLECHT = 'M'"
    sSQL = sSQL & " and Anrede = ''"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 20
    
    sSQL = "UpdATE Kunden set Anrede = 'Herr'"
    sSQL = sSQL & " where GESCHLECHT = 'M'"
    sSQL = sSQL & " and Anrede is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 21
    
    sSQL = "Delete from  Kunden"
    sSQL = sSQL & " where name is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 22
    
    sSQL = "UpdATE Kunden set KUERZEL = UCASE(LEFT(NAME,5))"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 23
 
    loeschNEW "EWWS0018", dbWK
        
    'Kunden ENDE
    
''''    'Artikelgruppen
''''
''''    lblx.Caption = "Artikelgruppen werden importiert..."
''''    lblx.Refresh
''''
''''    loeschNEW "EWWS0002", dbWK
''''    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0002"
''''
''''    loeschNEW "AGNDBFsic", dbWK
''''    sSQL = "Select * into AGNDBFsic from AGNDBF "
''''    dbWK.Execute sSQL, dbFailOnError
''''
''''    sSQL = "Delete from AGNDBF "
''''    dbWK.Execute sSQL, dbFailOnError
''''
''''
''''    sSQL = "Insert into AGNDBF Select  "
''''    sSQL = sSQL & " val(artgruppe) as AGN "
''''    sSQL = sSQL & ", text  as AGTEXT"
''''    sSQL = sSQL & ", abwertung as INVAB "
''''    sSQL = sSQL & " from EWWS0002 "
''''    dbWK.Execute sSQL, dbFailOnError
''''
''''    loeschNEW "EWWS0002", dbWK
        
    'Artikelgruppen ENDE
    
    'Firma
    txtStatus.Text = 24

    lblx.Caption = "Firma wird importiert..."
    lblx.Refresh

    loeschNEW "FIRMA", dbWK
    TransferTab gdBase, cPfad & "KissEWWS_" & iFil & ".mdb", "FIRMA"

    'Firma ENDE
    
    'SAP
    loeschNEW "SAP", dbWK
    CreateTableT2 "SAP", dbWK
    
    sSQL = "Insert into SAP (Datum) values ('" & DateValue(Now) + 60 & "')"
    dbWK.Execute sSQL, dbFailOnError
    'SAP ENDE
    
    'Bediener
    txtStatus.Text = 24
    
    lblx.Caption = "Bediener werden importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0025", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0025"
        
    sSQL = "Insert into Bedname Select  "
    sSQL = sSQL & " Nummer as bednu "
    sSQL = sSQL & ", vorName & ' ' & name as bedname"
    sSQL = sSQL & ", 'KISS' as passwort "
    sSQL = sSQL & ", 9 as bediener "
    sSQL = sSQL & " from EWWS0025 "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from bedname "
    sSQL = sSQL & " where bedname is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from bedname "
    sSQL = sSQL & " where bedname = '' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Bedname (bednu,bedName,Passwort,bediener) values ( 98,'KISS','KISS',9)  "
    dbWK.Execute sSQL, dbFailOnError
    
    loeschNEW "EWWS0025", dbWK
        
    'Bediener ENDE
    
    
    
    'Umsatz
    txtStatus.Text = 26
    
    lblx.Caption = "Umsatz werden importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0005", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0005"
        
    sSQL = "Insert into Umsatz Select  "
    sSQL = sSQL & " Datum "
    sSQL = sSQL & ", anzbon as KUNZ1 "
    sSQL = sSQL & ", MWSTBASIS1 as umsv1 "
    sSQL = sSQL & ", MWSTBASIS2 as umse1 "
    sSQL = sSQL & ", SATZ_1 as umsg1 "
    sSQL = sSQL & " from EWWS0005 where val(filiale) =  " & iFil
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Delete from Umsatz"
    sSQL = sSQL & " where Datum is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set kunz1 = 0"
    sSQL = sSQL & " where kunz1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set ekpr1 = 0"
    sSQL = sSQL & " where ekpr1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set kred1 = 0"
    sSQL = sSQL & " where kred1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set umsg1 = 0"
    sSQL = sSQL & " where umsg1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set umsv1 = 0"
    sSQL = sSQL & " where umsv1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set umse1 = 0"
    sSQL = sSQL & " where umse1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Umsatz set umso1 = 0"
    sSQL = sSQL & " where umso1 is null "
    dbWK.Execute sSQL, dbFailOnError
        
    'Umsatz ENDE
    
    If iBisFil > 0 Then
    
        'einmalig ZUMSATZ
        
        txtStatus.Text = 27
        
        
        sSQL = "Create table ZUmsatz "
        sSQL = sSQL & " ( "
        sSQL = sSQL & " Datum DateTime "
        sSQL = sSQL & ", UMSG double "
        sSQL = sSQL & ", UMSV double "
        sSQL = sSQL & ", UMSE double "
        sSQL = sSQL & ", UMSO double "
        sSQL = sSQL & ", Kunz long "
        sSQL = sSQL & ", EKPR double "
        sSQL = sSQL & ", Kred double "
        sSQL = sSQL & ", FILIALE long "
        sSQL = sSQL & " ) "
        dbWK.Execute sSQL, dbFailOnError
        
        For i = 1 To iBisFil
            
            sSQL = "Insert into ZUmsatz Select  "
            sSQL = sSQL & " Datum "
            sSQL = sSQL & ", anzbon as KUNZ "
            sSQL = sSQL & ", MWSTBASIS1 as umsv "
            sSQL = sSQL & ", MWSTBASIS2 as umse "
            sSQL = sSQL & ", SATZ_1 as umsg "
            sSQL = sSQL & ", " & i & " as filiale "
            sSQL = sSQL & " from EWWS0005 where val(filiale) =  " & i
            dbWK.Execute sSQL, dbFailOnError
        
        Next i
        
        
        sSQL = "Delete from  ZUmsatz"
        sSQL = sSQL & " where Datum is null "
        dbWK.Execute sSQL, dbFailOnError
        
        
        
        sSQL = "UpdATE ZUmsatz set kunz = 0"
        sSQL = sSQL & " where kunz is null "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE ZUmsatz set ekpr = 0"
        sSQL = sSQL & " where ekpr is null "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE ZUmsatz set kred = 0"
        sSQL = sSQL & " where kred is null "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE ZUmsatz set umsg = 0"
        sSQL = sSQL & " where umsg is null "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE ZUmsatz set umsv = 0"
        sSQL = sSQL & " where umsv is null "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE ZUmsatz set umse = 0"
        sSQL = sSQL & " where umse is null "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE ZUmsatz set umso = 0"
        sSQL = sSQL & " where umso is null "
        dbWK.Execute sSQL, dbFailOnError
        
        'ENDE einmalig ZUMSATZ
        
    End If
    
    loeschNEW "EWWS0005", dbWK
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Bondruck
    
    txtStatus.Text = 27
    
'''    lblx.Caption = "Bondruck werden importiert..."
'''    lblx.Refresh
'''
'''    loeschNEW "EWWS0008", dbWk
'''    TransferTab dbEWWS, cPfad & "KissEWWS_" & ifil & ".mdb", "EWWS0008"
'''
'''    sSQL = "Insert into BONTEXT Select  "
'''    sSQL = sSQL & " Bon1druck as zeilentext "
'''    sSQL = sSQL & ", 0 as zeilennr "
'''    sSQL = sSQL & " from EWWS0008 "
'''    dbWk.Execute sSQL, dbFailOnError
'''
'''    sSQL = "Insert into BONTEXT Select  "
'''    sSQL = sSQL & " Bon2druck as zeilentext "
'''    sSQL = sSQL & ", 1 as zeilennr "
'''    sSQL = sSQL & " from EWWS0008 "
'''    dbWk.Execute sSQL, dbFailOnError
'''
'''    sSQL = "Insert into BONTEXT Select  "
'''    sSQL = sSQL & " Bon3druck as zeilentext "
'''    sSQL = sSQL & ", 10 as zeilennr "
'''    sSQL = sSQL & " from EWWS0008 "
'''    dbWk.Execute sSQL, dbFailOnError
'''
'''    sSQL = "Insert into BONTEXT Select  "
'''    sSQL = sSQL & " Bon4druck as zeilentext "
'''    sSQL = sSQL & ", 11 as zeilennr "
'''    sSQL = sSQL & " from EWWS0008 "
'''    dbWk.Execute sSQL, dbFailOnError
'''
'''    loeschNEW "EWWS0008", dbWk
        
    'Bondruck ENDE
    
    'ARTIKEL
    
    txtStatus.Text = 28
    
    lblx.Caption = "Artikel werden importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0007", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0007"
    
    txtStatus.Text = 29
    lblx.Caption = "EAN wird importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0013", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0013"
    
    txtStatus.Text = 30
    lblx.Caption = "Bestände werden importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0019", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0019"
    
    lblx.Caption = "Artikel werden verarbeitet..."
    lblx.Refresh
        
    txtStatus.Text = 31
    
    sSQL = " Update EWWS0007 set artgruppe = '0' where artgruppe is null"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = " Update EWWS0007 set MINIMENGE = '0' where MINIMENGE is null"
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Insert into Artikel Select "
    sSQL = sSQL & " Bezeichner as BEZEICH "
    sSQL = sSQL & ", nummer as artnr "
    sSQL = sSQL & ", val(artgruppe) as agn "
    sSQL = sSQL & ", val(MINIMENGE) as minmen "
    sSQL = sSQL & ", EANNUMMER as EAN "
    sSQL = sSQL & ", KDNEKPREIS as LEKPR "
    sSQL = sSQL & ", MITTLEKPR as EKPR "
    sSQL = sSQL & ", KDNVKPREIS as VKPR "
    sSQL = sSQL & ", KDNVKPREIS as KVKPR1 "
    sSQL = sSQL & ", AUFNAHME as aufdat "
    sSQL = sSQL & ", trim(LIEFBESTNR) as libesnr "
    sSQL = sSQL & ", 1 as lpz "
    sSQL = sSQL & ", inhalt "
    sSQL = sSQL & ", Trim(mmmwst) as mwst "
    sSQL = sSQL & ", Trim(mmauswahl) as gefuehrt "
    sSQL = sSQL & ", Trim(mmvksperre) as Preisschu "
    sSQL = sSQL & ", liefnummer as linr "
    sSQL = sSQL & ", trim(ucase(groessnbez)) as inhaltbez "
    
    sSQL = sSQL & ", trim(MMRKZ) as AWM "
    sSQL = sSQL & ", Trim(MMPREISETI) &  Trim(MMREGALETI) as notizen "
    
    sSQL = sSQL & ", Right(KEINRAB,1) as RABATT_OK "
    sSQL = sSQL & ", Right(KEINBONUS,1) as BONUS_OK "
    
    sSQL = sSQL & " from EWWS0007 "
    
    'nur die aktiven
'    sSQL = sSQL & " where Trim(mmauswahl) = '2' "
    
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set etimerk = 'O' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set etimerk = 'S' where notizen = '21'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set etimerk = 'S' where notizen = '31'"
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Update Artikel set etimerk = 'R' where notizen = '13'"
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Update Artikel set etimerk = 'R' where notizen = '12'"
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Update Artikel set etimerk = 'B' where notizen = '23'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set etimerk = 'B' where notizen = '22'"
    dbWK.Execute sSQL, dbFailOnError

    sSQL = "Update Artikel set etimerk = 'B' where notizen = '33'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set etimerk = 'B' where notizen = '32'"
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    sSQL = "UpdATE Artikel set RKZ = 'N'"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'J'"
    sSQL = sSQL & " where AWM = '7' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'J'"
    sSQL = sSQL & " where AWM = '8' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'J'"
    sSQL = sSQL & " where AWM = '9' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RKZ = 'J'"
    sSQL = sSQL & " where AWM = '10' "
    dbWK.Execute sSQL, dbFailOnError
    
    
    sSQL = "UpdATE Artikel set BONUS_OK = 'J'"
    sSQL = sSQL & " where BONUS_OK = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set BONUS_OK = 'N'"
    sSQL = sSQL & " where BONUS_OK = '1' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'J'"
    sSQL = sSQL & " where RABATT_OK = '0' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set RABATT_OK = 'N'"
    sSQL = sSQL & " where RABATT_OK = '1' "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    lblx.Caption = "Warengruppentasten werden angelegt..."
    lblx.Refresh
    
    txtStatus.Text = 32
    
    sSQL = "Insert into Artikel (BEZEICH,LINR,ARTNR,UMS_OK,RABATT_OK,AWM,LPZ,RKZ,PGN,AGN,gefuehrt)  "
    sSQL = sSQL & " values ('DROGERIE',500000,500001,'J','J','0',0,'N',0,0,'J') "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Artikel (BEZEICH,LINR,ARTNR,UMS_OK,RABATT_OK,AWM,LPZ,RKZ,PGN,AGN,gefuehrt)  "
    sSQL = sSQL & " values ('PARFÜMERIE',500000,500002,'J','J','0',0,'N',0,0,'J') "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Artikel (BEZEICH,LINR,ARTNR,UMS_OK,RABATT_OK,AWM,LPZ,RKZ,PGN,AGN,gefuehrt)  "
    sSQL = sSQL & " values ('ACCESSOIRES',500000,500003,'J','J','0',0,'N',0,0,'J') "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into WARENGRU (BEZEICH,ARTNR,WGNR,FAKTOR,SGROESSE,bname)  "
    sSQL = sSQL & " values ('DROGERIE',500001,1,'+',10,'') "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into WARENGRU (BEZEICH,ARTNR,WGNR,FAKTOR,SGROESSE,bname)  "
    sSQL = sSQL & " values ('PARFÜMERIE',500002,2,'+',10,'') "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into WARENGRU (BEZEICH,ARTNR,WGNR,FAKTOR,SGROESSE,bname)  "
    sSQL = sSQL & " values ('ACCESSOIRES',500003,3,'+',10,'') "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 33
    
    Set rsrs = dbWK.OpenRecordset("Artikel")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!BEZEICH) Then
                rsrs.Edit
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, ",", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "'", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, ";", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "!", "")
                rsrs!BEZEICH = SwapStr(rsrs!BEZEICH, "*", "")
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    lblx.Caption = "Mehrwertsteuer wird aktualisiert..."
    lblx.Refresh
    
    txtStatus.Text = 34
    
    sSQL = "UpdATE Artikel set MWST = 'E'"
    sSQL = sSQL & " where MWST = '2' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set MWST = 'V'"
    sSQL = sSQL & " where MWST = '1' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set gefuehrt = 'J'"
    sSQL = sSQL & " where gefuehrt = '2' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set gefuehrt = 'N'"
    sSQL = sSQL & " where gefuehrt = '1' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set Preisschu = 'J'"
    sSQL = sSQL & " where Preisschu = '1' "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set Preisschu = 'N'"
    sSQL = sSQL & " where Preisschu = '2' "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    lblx.Caption = "Artikelgruppen werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set AGN = 0"
    sSQL = sSQL & " where AGN is null "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    txtStatus.Text = 36
    
    
    
    txtStatus.Text = 37
    
    lblx.Caption = "Lieferanten werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set LINR = 0"
    sSQL = sSQL & " where LINR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set PGN = 0"
    sSQL = sSQL & " where PGN is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 38
    
    lblx.Caption = "Farbmerkmale werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set AWM = '0'"
    dbWK.Execute sSQL, dbFailOnError
    
''    sSQL = "UpdATE Artikel set RKZ = 'N'"
''    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 39
    
    lblx.Caption = "Bezeichnungen werden aktualisiert..."
    lblx.Refresh
    
'    sSQL = "UpdATE Artikel set gefuehrt = 'J'"
'    dbWk.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set UMS_OK = 'J'"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 40
    
    lblx.Caption = "Produktlinien werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set BEZEICH = ''"
    sSQL = sSQL & " where BEZEICH is null "
    dbWK.Execute sSQL, dbFailOnError
    
'    sSQL = "UpdATE Artikel set RABATT_OK = 'J'"
'    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 41
    
    lblx.Caption = "Einkaufspreise werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set LEKPR = 0 "
    sSQL = sSQL & " where LEKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set EKPR = 0 "
    sSQL = sSQL & " where EKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 42
    
    lblx.Caption = "Kassenverkaufspreise werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set KVKPR1 = 0 "
    sSQL = sSQL & " where KVKPR1 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set VKPR = 0 "
    sSQL = sSQL & " where VKPR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 43
    
    lblx.Caption = "Bestellnummern werden aktualisiert..."
    lblx.Refresh
    
    sSQL = "UpdATE Artikel set LIBESNR = ''"
    sSQL = sSQL & " where LIBESNR is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set MWST = 'V'"
    sSQL = sSQL & " where MWST is null "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 44
    
    '*Pname
    sSQL = "UpdATE WKEINSTE set PNAME = 'TopCos'"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update WKEINSTE Set TABFAK = '1,25' "
    gdApp.Execute sSQL, dbFailOnError
    gdTabfak = 1.25
    '*Pname
    
    
    
    
    txtStatus.Text = 45
    
    sSQL = "UpdATE MWSTSATZ set VOLL = 19"
    sSQL = sSQL & " , ERM = 7 WHERE FurJahr=" & Year(Date)
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "EAN wird eingetragen..."
    lblx.Refresh
    txtStatus.Text = 46
    sSQL = "Delete from EWWS0013 where EWWS0013.EANNUMMER in ( Select ean from artikel) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from EWWS0013 where EWWS0013.EANNUMMER is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from EWWS0013 where EWWS0013.EANNUMMER ='' "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "EAN wird eingetragen (1)..."
    lblx.Refresh
    txtStatus.Text = 47
    sSQL = "Update Artikel inner join EWWS0013 on Artikel.artnr = val(EWWS0013.artNummer) "
    sSQL = sSQL & " set Artikel.EAN2 = EWWS0013.EANNUMMER "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "EAN wird eingetragen (2)..."
    lblx.Refresh
    txtStatus.Text = 48
    sSQL = "Delete from EWWS0013 where EWWS0013.EANNUMMER in ( Select ean2 from artikel) "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "EAN wird eingetragen (3)..."
    lblx.Refresh
    txtStatus.Text = 49
    sSQL = "Update Artikel inner join EWWS0013 on Artikel.artnr = val(EWWS0013.artNummer) "
    sSQL = sSQL & " set Artikel.EAN3 = EWWS0013.EANNUMMER "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "EAN wird eingetragen (4)..."
    lblx.Refresh
    txtStatus.Text = 50
    sSQL = "Delete from EWWS0013 where EWWS0013.EANNUMMER in ( Select ean3 from artikel) "
    dbWK.Execute sSQL, dbFailOnError
    
    
    lblx.Caption = "Bestände werden aktualisiert..."
    lblx.Refresh
    txtStatus.Text = 51
    sSQL = "Update Artikel "
    sSQL = sSQL & " set bestand = 0"
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 53
    sSQL = "Update Artikel "
    sSQL = sSQL & " set ean = '' where ean = '0000000000000' "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 55
    sSQL = "Update Artikel "
    sSQL = sSQL & " set ean2 = '' where ean2 = '0000000000000' "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 58
    sSQL = "Update Artikel "
    sSQL = sSQL & " set ean3 = '' where ean3 = '0000000000000' "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "EANs werden aktualisiert..."
    lblx.Refresh
    
    txtStatus.Text = 59
    sSQL = "UpdATE Artikel set EAN = ''"
    sSQL = sSQL & " where EAN is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set EAN2 = ''"
    sSQL = sSQL & " where EAN2 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "UpdATE Artikel set EAN3 = ''"
    sSQL = sSQL & " where EAN3 is null "
    dbWK.Execute sSQL, dbFailOnError
    
    lblx.Caption = "Bestände werden aktualisiert(1)..."
    lblx.Refresh
    
    txtStatus.Text = 60
    
    sSQL = "UpdATE EWWS0019 set bestand = 0 "
    sSQL = sSQL & " where bestand is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from  EWWS0019 where EWWS0019.artNummer is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel inner join EWWS0019 on Artikel.artnr = val(EWWS0019.artNummer) "
    sSQL = sSQL & " set Artikel.bestand = EWWS0019.bestand "
    sSQL = sSQL & " where EWWS0019.bestand > 0 "
    sSQL = sSQL & " and val(EWWS0019.Filiale) = " & iFil
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 60
    
    
    Dim cFil As String
    
    
    If iBisFil > 0 Then
    
'        sSQL = "Create table ZBESTAND "
'        sSQL = sSQL & " ( "
'        sSQL = sSQL & " FILIALNR long "
'        sSQL = sSQL & ", BESTAND long "
'        sSQL = sSQL & ", ARTNR long "
'        sSQL = sSQL & ", LINR long "
'        sSQL = sSQL & ", MINBEST long "
'        sSQL = sSQL & ", KVKPR1 double "
'        sSQL = sSQL & ", LASTDATE DateTime "
'        sSQL = sSQL & ", LASTTIME Text(10) "
'        sSQL = sSQL & ", SYNSTATUS Text(1) "
'        sSQL = sSQL & " ) "
'        dbWK.Execute sSQL, dbFailOnError

        sSQL = "Delete from ZBESTAND "
        dbWK.Execute sSQL, dbFailOnError
        
        For i = 1 To iBisFil
        
            cFil = i
            If Len(cFil) = 1 Then
                cFil = "0" & cFil
            End If
        
            sSQL = "Insert into ZBESTAND Select  "
            sSQL = sSQL & "  Artnr "
            sSQL = sSQL & ", 0  as LINR "
            sSQL = sSQL & ", 0  as Bestand "
            sSQL = sSQL & ", 0  as MINBEST "
            sSQL = sSQL & ", 0  as KVKPR1 "
            sSQL = sSQL & ", " & i & " as filialnr "
            sSQL = sSQL & " from Artikel "
            dbWK.Execute sSQL, dbFailOnError
            
            sSQL = "Update ZBESTAND inner join EWWS0019 on ZBESTAND.artnr = val(EWWS0019.artNummer) "
            sSQL = sSQL & " set ZBESTAND.bestand = EWWS0019.bestand "
            sSQL = sSQL & " where EWWS0019.bestand > 0 "
            sSQL = sSQL & " and EWWS0019.Filiale = '" & cFil & "'"
            sSQL = sSQL & " and ZBESTAND.filialnr = " & i
            dbWK.Execute sSQL, dbFailOnError
        Next i
    End If
    
    'Filialen machen wir auch fertig
    If iBisFil > 0 Then
    
'        sSQL = "Create table Filialen "
'        sSQL = sSQL & " ( "
'        sSQL = sSQL & " FILIALNR long "
'        sSQL = sSQL & ", FILIALNAME Text(35) "
'        sSQL = sSQL & ", LASTDATE DateTime "
'        sSQL = sSQL & ", LASTTIME Text(10) "
'        sSQL = sSQL & " ) "
'        dbWK.Execute sSQL, dbFailOnError

        sSQL = "Delete from Filialen "
        dbWK.Execute sSQL, dbFailOnError
        
        For i = 1 To iBisFil
            sSQL = "Insert into Filialen (FILIALNR,FILIALNAME) values "
            sSQL = sSQL & " ( " & i & ",  'Filiale' & '" & i & "' )"
            dbWK.Execute sSQL, dbFailOnError
        Next i
    End If
    
    txtStatus.Text = 62
    sSQL = "Delete from ARTIKEL where trim(Bezeich) = 'FREI FUER NEU' "
    dbWK.Execute sSQL, dbFailOnError
    
    loeschNEW "EWWS0019", dbWK

    loeschNEW "EWWS0013", dbWK
    
    loeschNEW "EWWS0007", dbWK
        
    'ARTIKEL ENDE
    
    'Lieferanten
    
    txtStatus.Text = 64
    
    lblx.Caption = "Lieferanten werden importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0021", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0021"
        
    sSQL = "Insert into LISRT Select "
    sSQL = sSQL & " Nummer as LINR "
    sSQL = sSQL & ", trim(Name1) as LIEFBEZ "
    sSQL = sSQL & ", strasse "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", Ort as stadt "
    sSQL = sSQL & ", telefon as tel "
    sSQL = sSQL & ", telefax as fax "
    sSQL = sSQL & ", KDNRLIEF as KUNDNR "
    sSQL = sSQL & ", Vertreter as ktext"
    sSQL = sSQL & ", BESTTEXT & ' ' & Bemerkung as Notiz "
    sSQL = sSQL & " from EWWS0021 "
    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 66
    
    sSQL = "UpdATE LISRT set KUERZEL = UCASE(LEFT(LIEFBEZ,5))"
    sSQL = sSQL & " where linr in (Select linr from Artikel) "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LISRT where LIEFBEZ = ''"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LISRT where LIEFBEZ is null"
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LISRT (LINR,LIEFBEZ,KUERZEL) values (500000,'TASTATURWARENGRUPPE','TASTA') "
    dbWK.Execute sSQL, dbFailOnError
    
    loeschNEW "EWWS0021", dbWK
        
    'Lieferanten ENDE
    
    
    'Kassjour
    
    lblx.Caption = "Kassjour wird importiert..."
    lblx.Refresh
    
    txtStatus.Text = 67
    loeschNEW "EWWS0017", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0017"
    
    sSQL = "Create Index satzart on EWWS0017(satzart) "
    dbWK.Execute sSQL, dbFailOnError
    
    
'    cFil = iFil
'    If Len(cFil) = 1 Then
'        cFil = "0" & cFil
'    End If
'
'    txtStatus.Text = 68
'    sSQL = "Insert into Kassjour Select  "
'    sSQL = sSQL & " artnummer as artnr "
'    sSQL = sSQL & ", menge "
'    sSQL = sSQL & ", (menge * preis) as vkpr "
'    sSQL = sSQL & ", Datum as adate "
'
'    sSQL = sSQL & ", " & iFil & " as Filiale "
'
'    sSQL = sSQL & ", 1 as KASNUM "
'    sSQL = sSQL & ", 'J' as UMS_OK "
'    sSQL = sSQL & ", BEZEICHNER as BEZEICH "
'    sSQL = sSQL & ", KUNDENNR as KUNDNR "
'    sSQL = sSQL & ", ZEIT as AZEIT "
'    sSQL = sSQL & ", trim(MMMWST) as MWST "
'    sSQL = sSQL & ", VERK as BEDIENER "
'    sSQL = sSQL & " from EWWS0017 "
'    sSQL = sSQL & " where trim(satzart) <> '7' "
'    sSQL = sSQL & " and Filiale = '" & cFil & "'"
'    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    loeschNEW "KassjourZ", dbWK
    
    sSQL = "Select * into KassjourZ "
    sSQL = sSQL & " from Kassjour where artnr = - 1 "
    dbWK.Execute sSQL, dbFailOnError
    
    
    If iBisFil > 0 Then
    
        For i = 1 To iBisFil
            cFil = i
            If Len(cFil) = 1 Then
                cFil = "0" & cFil
            End If
        
            sSQL = "Insert into KassjourZ Select  "
            sSQL = sSQL & " artnummer as artnr "
            sSQL = sSQL & ", menge "
            sSQL = sSQL & ", (menge * preis) as vkpr "
            sSQL = sSQL & ", Datum as adate "
            sSQL = sSQL & ", " & i & " as Filiale "
            sSQL = sSQL & ", 1 as KASNUM "
            sSQL = sSQL & ", 'J' as UMS_OK "
            sSQL = sSQL & ", BEZEICHNER as BEZEICH "
            sSQL = sSQL & ", KUNDENNR as KUNDNR "
            sSQL = sSQL & ", ZEIT as AZEIT "
            sSQL = sSQL & ", left(ZEIT,2) & right(zeit,2) as belegnr "
            sSQL = sSQL & ", trim(MMMWST) as MWST "
            sSQL = sSQL & ", VERK as BEDIENER "
            sSQL = sSQL & " from EWWS0017 "
'            sSQL = sSQL & " where trim(satzart) <> '7' "
'            sSQL = sSQL & " and Filiale = '" & cFil & "'"


            sSQL = sSQL & " where Filiale = '" & cFil & "'"
            dbWK.Execute sSQL, dbFailOnError
            
'            sSQL = "Delete * from EWWS0017 "
'            sSQL = sSQL & " where Filiale = '" & cFil & "'"
'            dbWK.Execute sSQL, dbFailOnError
        Next i
        
        sSQL = "Update KassjourZ set Preis = vkpr  "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "Update KassjourZ set vkpr = 0 "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE KassjourZ set MWST = 'E'"
        sSQL = sSQL & " where MWST = '2' "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "UpdATE KassjourZ set MWST = 'V'"
        sSQL = sSQL & " where MWST = '1' "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "Create Index Artnr on KassjourZ(Artnr) "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "Create Index Filiale on KassjourZ(Filiale) "
        dbWK.Execute sSQL, dbFailOnError
        
        loeschNEW "KassjourX", dbWK
        sSQL = "Select * into KassjourX "
        sSQL = sSQL & " from Kassjour where artnr = - 1 "
        dbWK.Execute sSQL, dbFailOnError
        
        Dim m As Integer
        
        For i = 1 To iBisFil
        
            loeschNEW "KassjourZ" & i, dbWK
    
            sSQL = "Select * into KassjourZ" & i
            sSQL = sSQL & " from Kassjourz where Filiale = " & i
            dbWK.Execute sSQL, dbFailOnError
            
            sSQL = "Create Index Artnr on KassjourZ" & i & "(Artnr) "
            dbWK.Execute sSQL, dbFailOnError
            
            sSQL = "Create Index adate on KassjourZ" & i & "(adate) "
            dbWK.Execute sSQL, dbFailOnError
            
            For m = 1 To 12
                sSQL = "Update KassjourZ" & i & " inner join Artikel on KassjourZ" & i & ".artnr = Artikel.artnr "
                sSQL = sSQL & " set KassjourZ" & i & ".agn = Artikel.agn "
                sSQL = sSQL & " , KassjourZ" & i & ".ekpr = Artikel.ekpr "
                sSQL = sSQL & " , KassjourZ" & i & ".VKPR = Artikel.VKPR "
                sSQL = sSQL & " , KassjourZ" & i & ".ean = Artikel.ean "
                sSQL = sSQL & " where month(adate) = " & m
                dbWK.Execute sSQL, dbFailOnError
            Next m
            
            sSQL = "Insert into KassjourX "
            sSQL = sSQL & "select * from KassjourZ" & i
            dbWK.Execute sSQL, dbFailOnError
            
            loeschNEW "KassjourZ" & i, dbWK
            
        Next i
    End If
    
    
    
    
    
    lblx.Caption = "Kassjour wird aktualisiert..."
    lblx.Refresh
    txtStatus.Text = 69
    
    
    sSQL = "Insert into Kassjour select * from kassjourz where filiale = " & iFil
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    loeschNEW "KUNDKASS", dbWK
    CreateTable "KUNDKASS", dbWK
    
    sSQL = "Insert into KUNDKASS "
    sSQL = sSQL & " select Filiale "
    sSQL = sSQL & " ,ADATE "
    sSQL = sSQL & " ,KUNDNR "
    sSQL = sSQL & " ,ARTNR "
    sSQL = sSQL & " ,PREIS "
    sSQL = sSQL & " ,MENGE "
    sSQL = sSQL & " ,VKPR "
    sSQL = sSQL & " ,BEDIENER as BEDNR"
    sSQL = sSQL & " from kassjourz where KUNDNR > 0 "
    dbWK.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Alter Table KassjourX drop zbonnr "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Alter Table KassjourX drop abok "
    dbWK.Execute sSQL, dbFailOnError
    
    
    loeschNEW "kassjourz", dbWK
    
    loeschNEW "EWWS0017", dbWK
        
    'Kassjour ENDE
    
    'Zugang
    
    lblx.Caption = "Zugänge werden importiert..."
    lblx.Refresh
    
    loeschNEW "EWWS0068", dbWK
    TransferTab dbEWWS, cPfad & "KissEWWS_" & iFil & ".mdb", "EWWS0068"
    
    sSQL = "Delete from ZUGANG "
    dbWK.Execute sSQL, dbFailOnError
    
'    sSQL = "Alter Table ZUGANG add  ARTNUMMER Long"
'    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from EWWS0068 "
    sSQL = sSQL & " where  artnr is null "
    dbWK.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from EWWS0068 "
    sSQL = sSQL & " where  artnr = '' "
    dbWK.Execute sSQL, dbFailOnError
    
        
    txtStatus.Text = 75
    
    CheckIndex "EWWS0068", "Datum", "", dbWK
    
    For i = Year(DateValue(Now)) To Year(DateValue(Now)) - 3 Step -1
    
        sSQL = "Insert into ZUGANG Select  "
        sSQL = sSQL & " artnr  "
'        sSQL = sSQL & " val(artnr) as ARTNUMMER "
        sSQL = sSQL & ", DATUM  as ADATE"
        sSQL = sSQL & ", Zeit as UHRZEIT "
        sSQL = sSQL & ", Menge as Bewegung "
        sSQL = sSQL & ", ALT as BESTANDALT "
        sSQL = sSQL & ", NEU as BESTANDNEU "
        sSQL = sSQL & " from EWWS0068 "
        sSQL = sSQL & " where year(datum) = " & i & " "
        dbWK.Execute sSQL, dbFailOnError
        
        sSQL = "Delete  "
        sSQL = sSQL & " from EWWS0068 "
        sSQL = sSQL & " where year(datum) = " & i & " "
        dbWK.Execute sSQL, dbFailOnError
        
        txtStatus.Text = CInt(txtStatus.Text) + 1
    Next i
    
'    sSQL = "Update ZUGANG set artnr = ARTNUMMER "
'    dbWK.Execute sSQL, dbFailOnError
    
'    sSQL = "Alter Table ZUGANG drop ARTNUMMER "
'    dbWK.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 87
    
    CheckIndex "ZUGANG", "artnr", "", dbWK
    CheckIndex "artikel", "artnr", "", dbWK
    
    sSQL = "Update ZUGANG inner join artikel on zugang.artnr = artikel.artnr "
    sSQL = sSQL & " set Zugang.bezeich = artikel.bezeich "
    sSQL = sSQL & " , Zugang.ean = artikel.ean "
    sSQL = sSQL & " , Zugang.linr = artikel.linr "
    sSQL = sSQL & " , Zugang.ekpr = artikel.ekpr "
    sSQL = sSQL & " , Zugang.libesnr = artikel.libesnr "
    dbWK.Execute sSQL, dbFailOnError
    
    loeschNEW "EWWS0068", dbWK
        
    'Zugang
    
    
    
    
    
    
    
    
    lblx.Caption = "Artlief wird erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 88
    ArtliefReinigenkomplett lblx, dbWK
    
    txtStatus.Text = 89
    
    lblx.Caption = "Datenbank wird kopiert..."
    lblx.Refresh
    
    dbWK.Close
    
    dbEWWS.Close
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "WKLEER\"
    
    cOldpath = cPfad
    cOldpath = ShortPath(cOldpath)
    cOldpath = cOldpath & "KissEWWS_" & iFil & ".mdb"
    
    cPfad = gcDBPfad      'dabapfad + WKLEER
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    txtStatus.Text = 90
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
    
    txtStatus.Text = 91
    
    gdBase.Close
    Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
    
    ReIndiziereArtikelWKL00 gdBase
'    db_Reindizieren gdBase, lblx, frmWKL151.txtStatus, frmWKL151.lbl6(28)
    
    lblx.Caption = "Artikelumsätze werden erstellt..."
    lblx.Refresh
    
    txtStatus.Text = 92
    
    UmsartjNew lblx
    
    txtStatus.Text = 93
    Ums_artNew lblx
    
    Dim sTabc As String
    sTabc = kassetabcheck(gdBase, lbl6(53), lbl6(28))
    
    If sTabc = "" Then

    Else
        MsgBox "Die Tabelle " & sTabc & " wurde nicht gefunden.", vbInformation, "Winkiss Hinweis:"
'                End
    End If
    
    lbl6(53).Caption = ""
    lbl6(28).Caption = ""

    txtStatus.Text = 100

    anzeige "Erfolg", "Fertig! Ihre Daten sind übernommen.", lblx
'    lblx.Caption = "Fertig! Ihre Daten sind übernommen."
'    lblx.Refresh
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "EWWSImport2"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
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
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Esüdro EWWS Import ist ein Fehler aufgetreten."
    
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
