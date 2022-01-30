VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL48 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Bestandsliste der verkauften Artikel"
   ClientHeight    =   8625
   ClientLeft      =   915
   ClientTop       =   480
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmWKL48.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   0
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   12015
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5100
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   11655
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   11655
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   2
         Left            =   10080
         TabIndex        =   20
         Top             =   5640
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Drucken"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Auswahlkriterien"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12015
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1215
         Left            =   4320
         TabIndex        =   9
         Top             =   120
         Width           =   5655
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Retouren"
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
            Index           =   3
            Left            =   3480
            TabIndex        =   13
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kollegenverkäufe"
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
            Index           =   2
            Left            =   3480
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "nicht umsatzrelevante Verkäufe"
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
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   3135
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0C000&
            Caption         =   "umsatzrelevante Verkäufe"
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
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "in Ergebnisliste berücksichtigen:"
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
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   2895
         End
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Heute"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   1
         Left            =   10080
         TabIndex        =   16
         Top             =   840
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Schließen"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   0
         Left            =   10080
         TabIndex        =   15
         Top             =   240
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Suchen"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   8
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "[Auswahllisten mit Taste F2]"
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
         Index           =   4
         Left            =   1440
         TabIndex        =   14
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "AGN"
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
         Index           =   3
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Linie"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant:"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Datum:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Bestandsliste der verkauften Artikel"
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
      TabIndex        =   22
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmWKL48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub DruckeErgebnisWKL48()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cTitel As String
    Dim cTitel2 As String
    Dim cLBSatz As String
    Dim lcount As Long
    
    loeschNEW "DRU_LISTE", gdBase
    
    cSQL = "Create Table DRU_LISTE "
    cSQL = cSQL & "(FELD1 Text(200), TITEL Text(200), TITEL2 Text(200) )"
    gdBase.Execute cSQL, dbFailOnError
    
    cTitel2 = "Artikelbestand nach Verkauf - Protokoll vom " & Format$(Now, "DD.MM.YYYY")
    cTitel = List1.list(0)
    cSQL = "Insert into DRU_LISTE ( TITEL, TITEL2 ) values ('" & cTitel & "', '" & cTitel2 & "' ) "
    gdBase.Execute cSQL, dbFailOnError
    
    For lcount = 0 To List2.ListCount - 1
        cLBSatz = List2.list(lcount)
        cSQL = "Insert into DRU_LISTE ( TITEL, TITEL2, FELD1 ) values ('" & cTitel & "', '" & cTitel2 & "', '" & cLBSatz & "' ) "
        gdBase.Execute cSQL, dbFailOnError
    Next lcount
    
    reportbildschirm "WKL034", "aWKL48"

Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeErgebnisWKL48"
        Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ErzeugeTempTabelleWKL48()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cSQL As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    loeschNEW "WKL048", gdBase
    CreateTable "WKL048", gdBase
    
    
    cSQL = "Create Index ARTNR on WKL048 (ARTNR)"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ErzeugeTempTabelleWKL48"
        Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Function fnPruefeDialogEingabenWKL48() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    Dim lWert As Long
    
    fnPruefeDialogEingabenWKL48 = 0
    
    cFeld = MaskEdBox1(0).Text
    If Not IsDate(cFeld) Then
        fnPruefeDialogEingabenWKL48 = 1
        Exit Function
    End If
    
    If Check1(0).Value = vbChecked Then
        lWert = lWert + 1
    End If
    If Check1(1).Value = vbChecked Then
        lWert = lWert + 2
    End If
    If Check1(2).Value = vbChecked Then
        lWert = lWert + 4
    End If
    If Check1(3).Value = vbChecked Then
        lWert = lWert + 8
    End If
    
    If lWert = 0 Then
        fnPruefeDialogEingabenWKL48 = 2
        Exit Function
    End If
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDialogEingabenWKL48"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub InitDialogWKL48()
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(0).Text = "__.__.____"
    MaskEdBox1(1).Text = "______"
    MaskEdBox1(2).Text = "___"
    MaskEdBox1(3).Text = "____"
    Check1(0).Value = vbChecked
    Check1(1).Value = vbChecked
    Check1(2).Value = vbChecked
    Check1(3).Value = vbChecked
    
    List1.Clear
    List2.Clear
    
    List1.AddItem "ArtNr. Artikelbezeichnung                  LiefNr Lin   AGN   Bewegung    Bestand"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InitDialogWKL48"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MoveDaten2DialogWKL48()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim lWert As Long
    Dim cFeld As String
    Dim cLBSatz As String
    
    List2.Clear
    
    cSQL = "Select * from WKL048 order by LINR, LPZ, BEZEICH"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cLBSatz = ""
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = 0
            End If
            cFeld = Trim$(Str$(lWert))
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            cFeld = cFeld & Space$(35 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!linr) Then
                lWert = rsrs!linr
            Else
                lWert = 0
            End If
            cFeld = Trim$(Str$(lWert))
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!LPZ) Then
                lWert = rsrs!LPZ
            Else
                lWert = 0
            End If
            cFeld = Trim$(Str$(lWert))
            cFeld = Space$(3 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "

            If Not IsNull(rsrs!AGN) Then
                lWert = rsrs!AGN
            Else
                lWert = 0
            End If
            cFeld = Trim$(Str$(lWert))
            cFeld = Space$(5 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "

            If Not IsNull(rsrs!BEWEGUNG) Then
                lWert = rsrs!BEWEGUNG
            Else
                lWert = 0
            End If
            cFeld = Trim$(Str$(lWert))
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "

            If Not IsNull(rsrs!BESTAND) Then
                lWert = rsrs!BESTAND
            Else
                lWert = 0
            End If
            cFeld = Trim$(Str$(lWert))
            cFeld = Space$(10 - Len(cFeld)) & cFeld
            cLBSatz = cLBSatz & cFeld & " "
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
            
        Loop
        Frame3.Visible = True
    Else
        MsgBox "Keine Daten gefunden!", vbInformation, "INFO"
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveDaten2DialogWKL48"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SucheDatenWKL48()
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    Dim lDatum As Long
    Dim lSchalter As Long
    Dim lSuchLiNr As Long
    Dim lSuchLPZ As Long
    Dim lSuchAGN As Long
    
    Dim rsrs As Recordset
    Dim rsArt As Recordset
    Dim rsZ As Recordset
    
    Dim sSQL1 As String
    
    Dim lartnr As Long
    Dim cBezeich As String
    Dim lMenge As Long
    Dim lBestand As Long
    Dim lLinr As Long
    Dim lLpz As Long
    Dim lagn As Long
    
    Dim cSQL As String
    Dim bTreffer As Boolean
    
    cFeld = MaskEdBox1(0).Text
    lDatum = DateValue(cFeld)
    
    If MaskEdBox1(1).Text <> "______" Then
        lSuchLiNr = Val(MaskEdBox1(1).Text)
    Else
        lSuchLiNr = -1
    End If
    
    If MaskEdBox1(2).Text <> "___" Then
        lSuchLPZ = Val(MaskEdBox1(2).Text)
    Else
        lSuchLPZ = -1
    End If
    
    If MaskEdBox1(3).Text <> "____" Then
        lSuchAGN = Val(MaskEdBox1(3).Text)
    Else
        lSuchAGN = -1
    End If
    
    If Check1(0).Value = vbChecked Then
        lSchalter = lSchalter + 1
    End If
    
    If Check1(1).Value = vbChecked Then
        lSchalter = lSchalter + 2
    End If
    
    If Check1(2).Value = vbChecked Then
        lSchalter = lSchalter + 4
    End If
    
    If Check1(3).Value = vbChecked Then
        lSchalter = lSchalter + 8
    End If
    
    
    '*********************************
    '* Hier geht's los!!!
    '*********************************
    
    cSQL = "Delete from WKL048"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    
    

    
    If lSchalter And 1 Then
        '*********************************
        '* Umsatzrelevante Verkäufe lesen
        '*********************************
        
        '****************************************
        '* aus Datenbestand nach Kassenabschluß
        '****************************************
        
        cSQL = "Select ARTNR, BEZEICH, MENGE from KASSJOUR "
        cSQL = cSQL & "where ADATE = " & Trim$(Str$(lDatum)) & " "
        cSQL = cSQL & "and UMS_OK <> 'N' "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    lartnr = rsrs!artnr
                Else
                    lartnr = -1
                End If
                If lartnr = 303220 Then
                    lartnr = lartnr
                End If
                If Not IsNull(rsrs!Menge) Then
                    lMenge = rsrs!Menge
                Else
                    lMenge = 0
                End If
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                
                sSQL1 = " select * from artikel where artnr = " & lartnr
                Set rsArt = gdBase.OpenRecordset(sSQL1)
                
                
                If Not rsArt.EOF Then
                    If Not IsNull(rsArt!linr) Then
                        lLinr = rsArt!linr
                    Else
                        lLinr = 0
                    End If
                    If Not IsNull(rsArt!LPZ) Then
                        lLpz = rsArt!LPZ
                    Else
                        lLpz = 0
                    End If
                    If Not IsNull(rsArt!AGN) Then
                        lagn = rsArt!AGN
                    Else
                        lagn = 0
                    End If
                    If Not IsNull(rsArt!BESTAND) Then
                        lBestand = rsArt!BESTAND
                    Else
                        lBestand = 0
                    End If
                Else
                    lLinr = 0
                    lLpz = 0
                    lagn = 0
                    lBestand = 0
                End If
                rsArt.Close: Set rsArt = Nothing
                
                bTreffer = True
                
                If lSuchLiNr <> -1 And lSuchLiNr <> lLinr Then
                    bTreffer = False
                End If
                If lSuchLPZ <> -1 And lSuchLPZ <> lLpz Then
                    bTreffer = False
                End If
                If lSuchAGN <> -1 And lSuchAGN <> lagn Then
                    bTreffer = False
                End If
                        
                If bTreffer Then
                    sSQL1 = " select * from WKL048 where artnr = " & lartnr
                    Set rsZ = gdBase.OpenRecordset(sSQL1)
                    
                    If Not rsZ.EOF Then
                        rsZ.Edit
                        If Not IsNull(rsZ!BEWEGUNG) Then
                            rsZ!BEWEGUNG = rsZ!BEWEGUNG - lMenge
                        Else
                            rsZ!BEWEGUNG = lMenge * (-1)
                        End If
                        rsZ.Update
                    Else
                        rsZ.AddNew
                        rsZ!artnr = lartnr
                        rsZ!BEZEICH = cBezeich
                        rsZ!linr = lLinr
                        rsZ!LPZ = lLpz
                        rsZ!AGN = lagn
                        rsZ!BEWEGUNG = lMenge * (-1)
                        rsZ!BESTAND = lBestand
                        rsZ.Update
                    End If
                    rsZ.Close: Set rsZ = Nothing
                End If
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
       
    End If
    
    If lSchalter And 2 Then
        '****************************************
        '* Nicht umsatzrelevante Verkäufe lesen
        '****************************************
        
        '****************************************
        '* aus Datenbestand nach Kassenabschluß
        '****************************************
        
        cSQL = "Select ARTNR, BEZEICH, MENGE from KASSJOUR "
        cSQL = cSQL & "where ADATE = " & Trim$(Str$(lDatum)) & " "
        cSQL = cSQL & "and UMS_OK = 'N' "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    lartnr = rsrs!artnr
                Else
                    lartnr = -1
                End If
                If lartnr = 303220 Then
                    lartnr = lartnr
                End If
                If Not IsNull(rsrs!Menge) Then
                    lMenge = rsrs!Menge
                Else
                    lMenge = 0
                End If
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                
                sSQL1 = " select * from artikel where artnr = " & lartnr
                Set rsArt = gdBase.OpenRecordset(sSQL1)
                
                
                If Not rsArt.EOF Then
                    If Not IsNull(rsArt!linr) Then
                        lLinr = rsArt!linr
                    Else
                        lLinr = 0
                    End If
                    If Not IsNull(rsArt!LPZ) Then
                        lLpz = rsArt!LPZ
                    Else
                        lLpz = 0
                    End If
                    If Not IsNull(rsArt!AGN) Then
                        lagn = rsArt!AGN
                    Else
                        lagn = 0
                    End If
                    If Not IsNull(rsArt!BESTAND) Then
                        lBestand = rsArt!BESTAND
                    Else
                        lBestand = 0
                    End If
                Else
                    lLinr = 0
                    lLpz = 0
                    lagn = 0
                    lBestand = 0
                End If
                rsArt.Close: Set rsArt = Nothing
                
                bTreffer = True
                
                If lSuchLiNr <> -1 And lSuchLiNr <> lLinr Then
                    bTreffer = False
                End If
                If lSuchLPZ <> -1 And lSuchLPZ <> lLpz Then
                    bTreffer = False
                End If
                If lSuchAGN <> -1 And lSuchAGN <> lagn Then
                    bTreffer = False
                End If
                        
                If bTreffer Then
                    sSQL1 = " select * from WKL048 where artnr = " & lartnr
                    Set rsZ = gdBase.OpenRecordset(sSQL1)
                    
                    If Not rsZ.EOF Then
                        rsZ.Edit
                        If Not IsNull(rsZ!BEWEGUNG) Then
                            rsZ!BEWEGUNG = rsZ!BEWEGUNG - lMenge
                        Else
                            rsZ!BEWEGUNG = lMenge * (-1)
                        End If
                        rsZ.Update
                    Else
                        rsZ.AddNew
                        rsZ!artnr = lartnr
                        rsZ!BEZEICH = cBezeich
                        rsZ!linr = lLinr
                        rsZ!LPZ = lLpz
                        rsZ!AGN = lagn
                        rsZ!BEWEGUNG = lMenge * (-1)
                        rsZ!BESTAND = lBestand
                        rsZ.Update
                    End If
                    rsZ.Close: Set rsZ = Nothing
                End If
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        
    End If
    
    If lSchalter And 4 Then
        '****************************************
        '* Kollegenverkäufe lesen
        '****************************************
        cSQL = "Select ARTNR, BEZEICH, MENGE from KOLLVERK "
        cSQL = cSQL & "where ADATE = " & Trim$(Str$(lDatum)) & " "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    lartnr = rsrs!artnr
                Else
                    lartnr = -1
                End If
                If lartnr = 303220 Then
                    lartnr = lartnr
                End If
                If Not IsNull(rsrs!Menge) Then
                    lMenge = rsrs!Menge
                Else
                    lMenge = 0
                End If
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                
                sSQL1 = " select * from artikel where artnr = " & lartnr
                Set rsArt = gdBase.OpenRecordset(sSQL1)
                
                
                If Not rsArt.EOF Then
                    If Not IsNull(rsArt!linr) Then
                        lLinr = rsArt!linr
                    Else
                        lLinr = 0
                    End If
                    If Not IsNull(rsArt!LPZ) Then
                        lLpz = rsArt!LPZ
                    Else
                        lLpz = 0
                    End If
                    If Not IsNull(rsArt!AGN) Then
                        lagn = rsArt!AGN
                    Else
                        lagn = 0
                    End If
                    If Not IsNull(rsArt!BESTAND) Then
                        lBestand = rsArt!BESTAND
                    Else
                        lBestand = 0
                    End If
                Else
                    lLinr = 0
                    lLpz = 0
                    lagn = 0
                    lBestand = 0
                End If
                rsArt.Close: Set rsArt = Nothing
                
                bTreffer = True
                
                If lSuchLiNr <> -1 And lSuchLiNr <> lLinr Then
                    bTreffer = False
                End If
                If lSuchLPZ <> -1 And lSuchLPZ <> lLpz Then
                    bTreffer = False
                End If
                If lSuchAGN <> -1 And lSuchAGN <> lagn Then
                    bTreffer = False
                End If
                        
                If bTreffer Then
                    sSQL1 = " select * from WKL048 where artnr = " & lartnr
                    Set rsZ = gdBase.OpenRecordset(sSQL1)
                    
                    If Not rsZ.EOF Then
                        rsZ.Edit
                        If Not IsNull(rsZ!BEWEGUNG) Then
                            rsZ!BEWEGUNG = rsZ!BEWEGUNG - lMenge
                        Else
                            rsZ!BEWEGUNG = lMenge * (-1)
                        End If
                        rsZ.Update
                    Else
                        rsZ.AddNew
                        rsZ!artnr = lartnr
                        rsZ!BEZEICH = cBezeich
                        rsZ!linr = lLinr
                        rsZ!LPZ = lLpz
                        rsZ!AGN = lagn
                        rsZ!BEWEGUNG = lMenge * (-1)
                        rsZ!BESTAND = lBestand
                        rsZ.Update
                    End If
                    rsZ.Close: Set rsZ = Nothing
                End If
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    If lSchalter And 8 Then
        '****************************************
        '* Retouren lesen
        '****************************************
        cSQL = "Select ARTNR, BEZEICH, MENGE from RETOURE "
        cSQL = cSQL & "where ADATE = " & Trim$(Str$(lDatum)) & " "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    lartnr = rsrs!artnr
                Else
                    lartnr = -1
                End If
                If lartnr = 303220 Then
                    lartnr = lartnr
                End If
                
                If Not IsNull(rsrs!Menge) Then
                    lMenge = rsrs!Menge
                Else
                    lMenge = 0
                End If
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                Else
                    cBezeich = ""
                End If
                
                 sSQL1 = " select * from artikel where artnr = " & lartnr
                Set rsArt = gdBase.OpenRecordset(sSQL1)
                
                
                If Not rsArt.EOF Then
                    If Not IsNull(rsArt!linr) Then
                        lLinr = rsArt!linr
                    Else
                        lLinr = 0
                    End If
                    If Not IsNull(rsArt!LPZ) Then
                        lLpz = rsArt!LPZ
                    Else
                        lLpz = 0
                    End If
                    If Not IsNull(rsArt!AGN) Then
                        lagn = rsArt!AGN
                    Else
                        lagn = 0
                    End If
                    If Not IsNull(rsArt!BESTAND) Then
                        lBestand = rsArt!BESTAND
                    Else
                        lBestand = 0
                    End If
                Else
                    lLinr = 0
                    lLpz = 0
                    lagn = 0
                    lBestand = 0
                End If
                rsArt.Close: Set rsArt = Nothing
                
                bTreffer = True
                
                If lSuchLiNr <> -1 And lSuchLiNr <> lLinr Then
                    bTreffer = False
                End If
                If lSuchLPZ <> -1 And lSuchLPZ <> lLpz Then
                    bTreffer = False
                End If
                If lSuchAGN <> -1 And lSuchAGN <> lagn Then
                    bTreffer = False
                End If
                        
                If bTreffer Then
                    sSQL1 = " select * from WKL048 where artnr = " & lartnr
                    Set rsZ = gdBase.OpenRecordset(sSQL1)
                    
                    If Not rsZ.EOF Then
                        rsZ.Edit
                        If Not IsNull(rsZ!BEWEGUNG) Then
                            rsZ!BEWEGUNG = rsZ!BEWEGUNG - lMenge
                        Else
                            rsZ!BEWEGUNG = lMenge * (-1)
                        End If
                        rsZ.Update
                    Else
                        rsZ.AddNew
                        rsZ!artnr = lartnr
                        rsZ!BEZEICH = cBezeich
                        rsZ!linr = lLinr
                        rsZ!LPZ = lLpz
                        rsZ!AGN = lagn
                        rsZ!BEWEGUNG = lMenge * (-1)
                        rsZ!BESTAND = lBestand
                        rsZ.Update
                    End If
                    rsZ.Close: Set rsZ = Nothing
                End If
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheDatenWKL48"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    positionierenwkl48
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    InitDialogWKL48
    
    ErzeugeTempTabelleWKL48
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub positionierenwkl48()
    On Error GoTo LOKAL_ERROR

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwkl48"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MaskEdBox1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(Index).BackColor = glSelBack1
    MaskEdBox1(Index).SelStart = 0
    MaskEdBox1(Index).SelLength = Len(MaskEdBox1(Index).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MaskEdBox1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld As String
    
    If KeyCode = vbKeyReturn Then
        SSCommand1_Click 0
    End If
    If KeyCode = vbKeyEscape Then
        SSCommand1_Click 1
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 0     'Datum
                'Kein Field-Prompt möglich
                
            Case Is = 1     'LiNr
                gF2Prompt.cFeld = "LINR"
            
            Case Is = 2     'LiNr
                If MaskEdBox1(1).Text = "______" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Exit Sub
                End If
                gF2Prompt.cFeld = "LPZ"
                gF2Prompt.cWert = Trim$(Str$(Val(MaskEdBox1(1).Text)))
            
            Case Is = 3     'AGN
                gF2Prompt.cFeld = "AGN"
                
        End Select
        
        If gF2Prompt.cFeld <> "" Then
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Select Case Index
                    Case Is = 0
                        'kein Field-Prompt möglich
                        
                    Case Is = 1     'LINR
                        cFeld = Trim$(Str$(Val(gF2Prompt.cWahl)))
                        cFeld = cFeld & String$(6 - Len(cFeld), "_")
                        
                    Case Is = 2     'LPZ
                        cFeld = Trim$(Str$(Val(gF2Prompt.cWahl)))
                        cFeld = cFeld & String$(3 - Len(cFeld), "_")
                        
                    Case Is = 3     'AGN
                        cFeld = Trim$(Str$(Val(gF2Prompt.cWahl)))
                        cFeld = cFeld & String$(3 - Len(cFeld), "_")
                End Select
                MaskEdBox1(Index).Text = cFeld
            End If
        End If
    End If
Exit Sub
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MaskEdBox1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    MaskEdBox1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MaskEdBox1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Dim lRet As Long
    
    Select Case Index
        Case Is = 0
            lRet = fnPruefeDialogEingabenWKL48()
            Select Case lRet
                Case Is = 0
                    SucheDatenWKL48
                    MoveDaten2DialogWKL48
                    
                Case Is = 1
                    MsgBox "Fehlendes oder falsches Datum!", vbCritical, "STOP!"
                    MaskEdBox1(0).SetFocus
                Case Is = 2
                    MsgBox "Bitte mindestens eine Verkaufstabelle auswählen!", vbCritical, "STOP!"
                    Check1(0).SetFocus
            End Select
            
            
        Case Is = 1
            Unload frmWKL48
            
            
        Case Is = 2
            DruckeErgebnisWKL48
            
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SSCommand2_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lHeute As Long
    Dim cHeute As String
    
    lHeute = Fix(Now)
    cHeute = Format$(lHeute, "DD.MM.YYYY")
    MaskEdBox1(0).Text = cHeute
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil Bestandsliste verkaufter Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


