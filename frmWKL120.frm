VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL120 
   Caption         =   "Kundenimport"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL120.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox Check1 
      Caption         =   "nur mit Emailadresse"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2655
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   2295
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   1
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
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2295
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
      Caption         =   "Excel Import"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   2295
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
   Begin VB.Label Label1 
      Caption         =   "für Outlook Kunden als Kontakte exportieren (csv - Datei)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   12
      Top             =   5400
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   $"frmWKL120.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   11295
   End
   Begin VB.Label Label1 
      Caption         =   $"frmWKL120.frx":04D8
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   11295
   End
   Begin VB.Label Label1 
      Caption         =   "für Outlook Kunden als Kontakte exportieren (mdb - Access Datei)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   4800
      Width           =   8895
   End
   Begin VB.Label Label26 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   11535
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   3
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
      Caption         =   "Kundenimport / Export"
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
      Width           =   9135
   End
End
Attribute VB_Name = "frmWKL120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            Unload frmWKL120
        Case 1 'excel import
            Excelimport
        Case 2
            outlookKundenKontakte
        Case 3
            Export_inCSV_OutlookKontakte
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenimport/Export ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub outlookKundenKontakte()
On Error GoTo LOKAL_ERROR

    Dim dabaStat    As Database
    Dim cPfad       As String
    Dim cSQL        As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "BOX\"

    Kill cPfad & "Kontakte.mdb"
    
    loeschNEW "KONTAKTE", gdBase
    CreateTableT2 "KONTAKTE", gdBase
    
    cSQL = cSQL & "Insert into kontakte Select Anrede "
    cSQL = cSQL & ", VORNAME "
    cSQL = cSQL & ", Name as NACHNAME "
    cSQL = cSQL & ", FIRMA "
    cSQL = cSQL & ", Strasse as STRAßEPRIVAT "
    cSQL = cSQL & ", Stadt as ORTPRIVAT "
    cSQL = cSQL & ", PLZ as POSTLEITZAHLPRIVAT "
    cSQL = cSQL & ", TEL as TELEFONPRIVAT "
    cSQL = cSQL & ", FAXNR as FAXPRIVAT "
    cSQL = cSQL & ", MOBILTEL as MOBILTELEFON "
    cSQL = cSQL & ", EMAIL  as [E-Mail]"
'    cSQL = cSQL & ", EMAIL as EMAILANGEZEIGTERNAME2 "
    cSQL = cSQL & ", Datum1 as GEBURTSTAG"
    cSQL = cSQL & ", GESCHLECHT "
    cSQL = cSQL & ", 'Normal' as PRIORITÄT "
    cSQL = cSQL & ", 'Normal' as VERTRAULICHKEIT "
    cSQL = cSQL & ", Kundnr as Position"
    cSQL = cSQL & ", 0 as PRIVAT"
    cSQL = cSQL & " from Kunden "
'    cSQL = cSQL & " where name like 'A*' "
    If Check1.Value = vbChecked Then
        cSQL = cSQL & " where email <> '' "
    End If
    gdBase.Execute cSQL, dbFailOnError
        
    Set dabaStat = CreateDatabase(cPfad & "Kontakte.mdb", dbLangGeneral, dbVersion40)
    TransferTab gdBase, cPfad & "Kontakte.mdb", "Kontakte"
    dabaStat.Close
    
    

    MsgBox "Diese Datei ist unter ('" & cPfad & "') mit dem Namen: 'Kontakte.mdb' abgespeichert", vbInformation, "Winkiss Information:"
                
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "outlookKundenKontakte"
        Fehler.gsFehlertext = "Im Programmteil Kundenimport/Export ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Export_inCSV_OutlookKontakte()
   On Error GoTo LOKAL_ERROR

    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim cSatz       As String
    Dim rsrs        As Recordset
    Dim cSQL        As String
    Dim cPfad       As String
    Dim cdatei      As String
   
    cdatei = "Kontakte.csv"
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "BOX\"

    Kill cPfad & cdatei
    
    iFileNr = FreeFile
    Open cPfad & cdatei For Binary As #iFileNr
    
    cSatz = "NACHNAME,VORNAME,E-MAIL,STRAßE PRIVAT,POSTLEITZAHL PRIVAT,ORT PRIVAT,TELEFON (PRIVAT),FAX PRIVAT,MOBILTELEFON,GEBURTSTAG" & vbCrLf
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    cSQL = "Select * from Kunden "
    If Check1.Value = vbChecked Then
        cSQL = cSQL & " where email <> '' "
    End If
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            
                cSatz = ""
'                If Not IsNull(rsrs!anrede) Then
'                    cSatz = cSatz & rsrs!anrede & ";"
'                Else
'                    cSatz = cSatz & ";"
'                End If

                If Not IsNull(rsrs!name) Then
                    cSatz = cSatz & rsrs!name & ","
                Else
                    cSatz = cSatz & ""
                End If
                
                If Not IsNull(rsrs!vorname) Then
                    cSatz = cSatz & rsrs!vorname & ","
                Else
                    cSatz = cSatz & ","
                End If
                
                If Not IsNull(rsrs!Email) Then
                    cSatz = cSatz & rsrs!Email & ","
                Else
                    cSatz = cSatz & ","
                End If
                
                If Not IsNull(rsrs!strasse) Then
                    cSatz = cSatz & rsrs!strasse & ","
                Else
                    cSatz = cSatz & ","
                End If
                
                If Not IsNull(rsrs!Plz) Then
                    cSatz = cSatz & rsrs!Plz & ","
                Else
                    cSatz = cSatz & ","
                End If
                
                If Not IsNull(rsrs!STADT) Then
                    cSatz = cSatz & rsrs!STADT & ","
                Else
                    cSatz = cSatz & ","
                End If
                
                 If Not IsNull(rsrs!Tel) Then
                    cSatz = cSatz & rsrs!Tel & ","
                Else
                    cSatz = cSatz & ","
                End If

                If Not IsNull(rsrs!FAXNR) Then
                    cSatz = cSatz & rsrs!FAXNR & ","
                Else
                    cSatz = cSatz & ","
                End If
                
                 If Not IsNull(rsrs!Mobiltel) Then
                    cSatz = cSatz & rsrs!Mobiltel & ","
                Else
                    cSatz = cSatz & ","
                End If

                If Not IsNull(rsrs!Datum1) Then
                    cSatz = cSatz & rsrs!Datum1 & vbCrLf
                Else
                    cSatz = cSatz & vbCrLf
                End If
                
                
                
'                If Not IsNull(rsrs!firma) Then
'                    cSatz = cSatz & rsrs!firma & ";"
'                Else
'                    cSatz = cSatz & ";"
'                End If
'
'                If Not IsNull(rsrs!geschlecht) Then
'                    cSatz = cSatz & rsrs!geschlecht & ";"
'                Else
'                    cSatz = cSatz & ";"
'                End If
'
'
'                cSatz = cSatz & "Normal" & ";"
'                cSatz = cSatz & "Normal" & ";"
'
'                If Not IsNull(rsrs!Kundnr) Then
'                    cSatz = cSatz & rsrs!Kundnr & ";"
'                Else
'                    cSatz = cSatz & ";"
'                End If
'
'                cSatz = cSatz & "0" & vbCrLf
                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cSatz
                
           
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close
    Close iFileNr
    
    MsgBox "Diese Datei ist unter ('" & cPfad & "') mit dem Namen: 'Kontakte.csv' abgespeichert", vbInformation, "Winkiss Information:"

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Export_inCSV_OutlookKontakte"
        Fehler.gsFehlertext = "Im Programmteil Kundenimport/Export ist ein Fehler aufgetreten."
        Fehlermeldung1
    End If
End Sub
Private Sub Excelimport()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cPfad As String
    Dim cDatname As String
    Dim dbExcel As Database
    Dim lAnz As Long
    Dim lAnzF As Long
    Dim lAnzZ As Long
    Dim rsrs As Recordset
    Dim rsKU As Recordset
    Dim gsExcel50 As String
    Dim bFound  As Boolean
    Dim cFehlerInhalt As String
    gsExcel50 = "Excel 5.0;"
    
    If pfadseekExcelkuim = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label1(4)
        Exit Sub
    End If
    
    Screen.MousePointer = 11

    anzeige "normal", "", Label1(4)
    cPfad = Label1(0).Caption
    
    Set dbExcel = OpenDatabase(cPfad, 0, 0, gsExcel50)

    bFound = False
    Dim sVergebneKDNR As String
    
    Set rsrs = dbExcel.OpenRecordset("Neu$")
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ID) Then
                If fnIsKundenNrfrei(CLng(rsrs!ID)) Then
                    
                Else
                    sVergebneKDNR = rsrs!ID
                    
                    bFound = True
                    Exit Do
                End If
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If bFound = False Then
    
        Set rsKU = gdBase.OpenRecordset("Kunden")
    
        lAnzZ = 0
        Set rsrs = dbExcel.OpenRecordset("Neu$")
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!ID) Then
                    rsKU.AddNew
                    
                    lAnzZ = lAnzZ + 1
                    
                    rsKU!Kundnr = rsrs!ID
                    rsKU!Kuerzel = UCase(Left(rsrs!name, 5))
                    rsKU!firma = rsrs!firma
                    rsKU!titel = rsrs!titel
                    rsKU!name = rsrs!name
                    rsKU!vorname = rsrs!vorname
                    rsKU!strasse = Left(rsrs!straße, 35)
                    rsKU!Plz = rsrs!Plz
                    rsKU!STADT = rsrs!Ort
                    rsKU!Tel = rsrs!telefon
                    
                    If Not IsNull(rsrs!TeleFAX) Then
                        rsKU!FAXNR = rsrs!TeleFAX
                    End If
                    
                    If Not IsNull(rsrs!Mobiltel) Then
                        rsKU!Mobiltel = Left(rsrs!Mobiltel, 15)
                    End If
                    
                    
                    rsKU!KUNDKART = rsrs!PAYBACK
                    rsKU!Email = rsrs!Email
                    rsKU!Datum1 = rsrs!GEBDATUM
                    
                    rsKU!anrede = rsrs!anrede
                    
                    If Not IsNull(rsrs!Makeup) Then
                        rsKU!NOTIZEN = "Make up " & rsrs!Makeup & " "
                    End If
                    
                    If Not IsNull(rsrs!Pflege) Then
                        rsKU!NOTIZEN = rsKU!NOTIZEN & "Pflege " & rsrs!Pflege & " "
                    End If
                    
                    If Not IsNull(rsrs!Zusatzinfo) Then
                        rsKU!NOTIZEN = rsKU!NOTIZEN & "Zusatzinfo " & rsrs!Zusatzinfo & " "
                    End If
                    rsKU!angelegt = DateValue(Now)
                    
                    rsKU!FILIALNR = 0
                    If Trim(UCase(rsrs!anrede)) = "FRAU" Then
                        rsKU!geschlecht = "W"
                    ElseIf Trim(UCase(rsrs!anrede)) = "HERR" Then
                        rsKU!geschlecht = "M"
                    End If
                    
                    rsKU!BONUS = 0
                    rsKU!RABATT = 0
                    rsKU!Status = "N"
                
                    rsKU.Update
                End If
            rsrs.MoveNext
            Loop
            
        End If
        rsrs.Close: Set rsrs = Nothing
        
        anzeige "normal", lAnzZ & " Kunden wurden korrekt eingelesen.", Label1(4)
    Else
        anzeige "rot", "(schon enthalten: " & sVergebneKDNR & ") Die Kunden können nicht eingelesen werden. Es sind schon einige Kundennummern vergeben.", Label1(4)
    End If
    
    
    
    Screen.MousePointer = 0

    dbExcel.Close
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3163 Then
        MsgBox rsrs!Kundnr & " Bei diesem Kunden konnte nicht alles übernommen werden. Bitte überprüfen Sie diesen nochmals.", vbCritical, "Winkiss Hinweis"
        lAnzF = lAnzF - 1
        Resume Next
    ElseIf err.Number = 3125 Then
        anzeige "rot", "Die Excelliste hat nicht das erwartete Format", Label1(4)
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Excelimport"
        Fehler.gsFehlertext = "Im Programmteil Kundenimport/Export ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Function pfadseekExcelkuim() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sExcelpfad  As String
    
    pfadseekExcelkuim = False

    sTitle = "Speichern des Pfades"
    
    sFilter = "Excel - Dateien (*.xls)|*.xls"
    
    sOldpfad = ""
    sExcelpfad = pfadaendernKomplett(sTitle, sFilter, sOldpfad)
    
    If sExcelpfad <> "" Then
        pfadseekExcelkuim = True
        Label1(0).Caption = sExcelpfad
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcelkuim"
    Fehler.gsFehlertext = "Im Programmteil Kundenimport/Export ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim sAnzeigetext As String

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    sAnzeigetext = "Möchten Sie Kundendaten importieren?" & vbCrLf
    sAnzeigetext = sAnzeigetext & "Stimmt das Format nicht überein? Rufen Sie uns an!" & vbCrLf
    sAnzeigetext = sAnzeigetext & "Wir passen Ihre Daten zum Importieren an." & vbCrLf
    sAnzeigetext = sAnzeigetext & "Telefon: 0511/9559112 " & vbCrLf
    
    Label26(1).Caption = sAnzeigetext
    Label26(1).Refresh
    
    anzeige "normal", "", Label1(4)
    Screen.MousePointer = 0
       
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenimport/Export ist ein Fehler aufgetreten."
    
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



