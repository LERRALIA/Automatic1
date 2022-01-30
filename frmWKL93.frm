VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL93 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "MDE betanken"
   ClientHeight    =   8625
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   11910
   Icon            =   "frmWKL93.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Auswahl-Listen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3480
      TabIndex        =   13
      Top             =   0
      Width           =   8295
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5280
         MaxLength       =   8
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox List5 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   7560
         Pattern         =   "*.LST"
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   5
         Left            =   5280
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Speichern"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   4
         Left            =   3480
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Auswählen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   ".LST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   7080
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Auswahlliste speichern unter:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   20
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   11775
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   4800
         TabIndex        =   10
         Top             =   720
         Width           =   4575
      End
      Begin VB.ListBox List3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4800
         TabIndex        =   9
         Top             =   480
         Width           =   4575
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   4575
      End
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   4575
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   3
         Left            =   9480
         TabIndex        =   17
         Top             =   5520
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Schließen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   14
         Top             =   4920
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Leeren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   12
         Top             =   480
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Fülle Ziel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   6
         Left            =   9480
         TabIndex        =   23
         Top             =   1080
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Übertragen an MDE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Font3D          =   3
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "ausgewählte Lieferanten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "vorhandene Lieferanten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Zieldatei"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         MaxLength       =   8
         TabIndex        =   2
         Text            =   "TO_MDE"
         Top             =   600
         Width           =   1695
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Ziel-Datei leeren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   ".DAT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Name der Datei für das MDE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmWKL93"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FuelleZielDateiMDE20()
    On Error GoTo LOKAL_ERROR
    
    Dim cLbSatz As String
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim cEAN As String
    Dim iFileNr As Integer
    Dim clinr As String
    Dim cLiNrMerker As String
    Dim cArtNr As String
    Dim lPos As Long
    
    Set rsrs = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    rsrs.Index = "LINR"
    
    iFileNr = FreeFile
        
    Open App.Path & "\TO_MDE.DAT" For Binary As #iFileNr
    
    lAnzSatz = List4.ListCount
    For lAktSatz = 0 To lAnzSatz - 1
        cLbSatz = List4.list(lAktSatz)
        If Len(cLbSatz) > 6 Then
            cLbSatz = Left$(cLbSatz, 6)
        End If
        cLbSatz = Trim$(cLbSatz)
        
        clinr = cLbSatz
        cLiNrMerker = cLbSatz
        
        rsrs.Seek "=", cLbSatz
        
        If rsrs.NoMatch Then
            'Nix tun!
        Else
            Do While clinr = cLiNrMerker
                If Not IsNull(rsrs!artnr) Then
                    cArtNr = rsrs!artnr
                Else
                    cArtNr = ""
                End If
            
                '***** Feld EAN *****
                If Not IsNull(rsrs!EAN) Then
                    cEAN = rsrs!EAN
                Else
                    cEAN = ""
                End If
                cEAN = Trim$(cEAN)
                If cEAN = "" Then
                    cEAN = cArtNr
                    cEAN = fnErzeugeEANCodeMDE(cEAN)
                End If
                cEAN = cEAN & vbCrLf
                                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cEAN
    
    
                '***** Feld EAN2 *****
    
                If Not IsNull(rsrs!EAN2) Then
                    cEAN = rsrs!EAN2
                Else
                    cEAN = ""
                End If
                cEAN = Trim$(cEAN)
                If cEAN <> "" Then
                    cEAN = cEAN & vbCrLf
                                    
                    lPos = LOF(iFileNr)
                    lPos = lPos + 1
                    Put #iFileNr, lPos, cEAN
                End If
            
                '***** Feld EAN3 *****
    
                If Not IsNull(rsrs!EAN3) Then
                    cEAN = rsrs!EAN3
                Else
                    cEAN = ""
                End If
                cEAN = Trim$(cEAN)
                If cEAN <> "" Then
                    cEAN = cEAN & vbCrLf
                                    
                    lPos = LOF(iFileNr)
                    lPos = lPos + 1
                    Put #iFileNr, lPos, cEAN
                End If
                
                rsrs.MoveNext
                    
                If Not IsNull(rsrs!linr) Then
                    clinr = rsrs!linr
                Else
                    clinr = "-1"
                End If
            
            Loop
            
        End If
    Next lAktSatz
    rsrs.Close: Set rsrs = Nothing
    
    lAnzSatz = LOF(iFileNr) / 14
    
    lAnzSatz = lAnzSatz / 100
    lAnzSatz = lAnzSatz * 100
    
    Close iFileNr
    
    If lAnzSatz > 30000 Then
        cSQL = "Datei steht für Übertragung an MDE bereit!" & vbCrLf & vbCrLf
        cSQL = cSQL & "HINWEIS:" & vbCrLf
        cSQL = cSQL & "Die Datei enthält ca. " & Format$(lAnzSatz, "###,##0") & " Datensätze." & vbCrLf
        cSQL = cSQL & vbCrLf
        cSQL = cSQL & "Bitte beachten Sie die maximale Aufnahmekapazität Ihres MDE-Gerätes!"
        
        SSCommand1(6).Enabled = True
        MsgBox cSQL, vbInformation, "Winkiss Hinweis:"
        
    Else
        SSCommand1(6).Enabled = True
        MsgBox "Datei steht für Übertragung an MDE bereit!", vbInformation, "Winkiss Hinweis:"
    End If
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuelleZielDateiMDE20"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function fnErzeugeEANCodeMDE(cCode As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim iCount As Integer
    Dim cZeichen As String
    Dim iWert As Integer
    Dim iSumme As Integer
    
    fnErzeugeEANCodeMDE = ""
    
    cCode = String$(6 - Len(cCode), "0") & cCode
    cCode = "2" & cCode
    
    iSumme = 0
    For iCount = 1 To 7
        cZeichen = Mid$(cCode, iCount, 1)
        iWert = Val(cZeichen)
        If iCount / 2 = Int(iCount / 2) Then
            iSumme = iSumme + iWert
        Else
            iSumme = iSumme + (iWert * 3)
        End If
    Next iCount
    
    iSumme = iSumme Mod 10
    If iSumme <> 0 Then
        iSumme = 10 - iSumme
    End If
    
    cZeichen = Trim$(Str$(iSumme))
    cCode = cCode & cZeichen
    
    fnErzeugeEANCodeMDE = cCode
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnErzeugeEANCodeMDE"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub LeseAuswahlListeMDE20()
    On Error GoTo LOKAL_ERROR
    
    Dim cLbSatz As String
    Dim iFileNr As Integer
    Dim cdatei As String
    Dim lStart As Long
    Dim lEnde As Long
    
    If List5.ListIndex < 0 Then
        MsgBox "Bitte einen Eintrag in der Liste auswählen!", vbCritical, "Winkiss Hinweis:"
        List5.SetFocus
        Exit Sub
    End If
    
    List4.Clear
    
    cLbSatz = List5.list(List5.ListIndex)
    
    iFileNr = FreeFile
    
    Open App.Path & "\" & cLbSatz For Binary As #iFileNr
    cdatei = Space$(LOF(iFileNr))
    Get #iFileNr, 1, cdatei
    Close iFileNr
    
    lStart = 1
    Do While lStart < Len(cdatei)
        lEnde = InStr(lStart, cdatei, vbCrLf)
        cLbSatz = Mid$(cdatei, lStart, lEnde - lStart)
        cLbSatz = Trim$(cLbSatz)
        List4.AddItem cLbSatz
        lStart = lEnde + 2
    Loop
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseAuswahlListeMDE20"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub LeseLieferantenMDE20()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim cLbSatz As String
    
    List2.Clear
    
    cSQL = "Select * from Lisrt order by KUERZEL"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!linr) Then
                cFeld = rsrs!linr
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = Space$(6 - Len(cFeld)) & cFeld
            cLbSatz = cFeld & " "
            
            If Not IsNull(rsrs!KUERZEL) Then
                cFeld = rsrs!KUERZEL
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(5 - Len(cFeld))
            cLbSatz = cLbSatz & cFeld & " "
            
            If Not IsNull(rsrs!LIEFBEZ) Then
                cFeld = rsrs!LIEFBEZ
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cLbSatz = cLbSatz & cFeld
            
            List2.AddItem cLbSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseLieferantenMDE20"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub LeseVorgabeListenMDE20()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    
    File1.Path = App.Path
    File1.Refresh
    
    List5.Clear
    
    lAnzSatz = File1.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        List5.AddItem UCase$(File1.list(lAktSatz))
    Next lAktSatz
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseVorgabeListenMDE20"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   

End Sub

Private Sub SpeichereAuswahlListeMDE20()
    On Error GoTo LOKAL_ERROR

    Dim cDateiName As String
    Dim iFileNr As Integer
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim cLbSatz As String
    Dim lPos As Long
    Dim iRet As Integer
    
    cDateiName = Text2.Text
    cDateiName = Trim$(cDateiName)
    If cDateiName = "" Then
        MsgBox "Bitte den Dateinamen angeben!", vbCritical, "Winkiss Hinweis:"
        Text2.SetFocus
        Exit Sub
    End If
    
    If List4.ListCount = 0 Then
        MsgBox "Bitte mindestens einen Lieferanten angeben!", vbCritical, "Winkiss Hinweis:"
        List2.SetFocus
        Exit Sub
    End If

    iFileNr = FreeFile
    
    Open App.Path & "\" & cDateiName & ".LST" For Binary As #iFileNr
    
    If LOF(iFileNr) > 0 Then
        iRet = MsgBox("Die Datei existiert bereits! Datei überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            Close iFileNr
            Kill App.Path & "\" & cDateiName & ".LST"
            Open App.Path & "\" & cDateiName & ".LST" For Binary As #iFileNr
        End If
    End If
    
    lAnzSatz = List4.ListCount
    
    For lAktSatz = 0 To lAnzSatz - 1
        cLbSatz = List4.list(lAktSatz)
        cLbSatz = cLbSatz & vbCrLf
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cLbSatz
    Next lAktSatz
    
    Close iFileNr
    
    Text2.Text = ""
    
    LeseVorgabeListenMDE20
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeichereAuswahllisteMDE20"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    
'    PositionierenWKL90
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    List1.Clear
    List1.AddItem "LiefNr Kurz  Bezeichnung"
    
    List2.Clear
    
    List3.Clear
    List3.AddItem "LiefNr Kurz  Bezeichnung"
    
    List4.Clear
    
    List5.Clear
    
    LeseVorgabeListenMDE20
    
    LeseLieferantenMDE20
    
    Text2.Text = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub List2_dblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim cLbSatz2 As String
    Dim cLbSatz4 As String
    Dim lAnzSatz As Long
    Dim lAktSatz As Long
    Dim bGefunden As Boolean
    
    cLbSatz2 = List2.list(List2.ListIndex)
    
    lAnzSatz = List4.ListCount
    bGefunden = False
    For lAktSatz = 0 To lAnzSatz - 1
        cLbSatz4 = List4.list(lAktSatz)
        If Trim$(cLbSatz4) = Trim$(cLbSatz2) Then
            bGefunden = True
            Exit For
        End If
    Next lAktSatz
    
    If Not bGefunden Then
        List4.AddItem cLbSatz2
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    

End Sub


Private Sub List4_Click()
    On Error GoTo LOKAL_ERROR
    
    List4.RemoveItem List4.ListIndex
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List4_Click"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub List5_DblClick()
On Error GoTo LOKAL_ERROR

    SSCommand1_Click 4
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List5_DblClick"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim i As Integer
    Dim iFileNr As Integer
    Screen.MousePointer = vbHourglass
    
    Select Case Index
    
        Case Is = 0     'Ziel-Datei leeren
            Kill App.Path & "\TO_MDE.DAT"
            MsgBox "To_Mde.dat erfolgreich gelöscht!", vbOKOnly + vbInformation, "Winkiss Hinweis:"
            
        Case Is = 1     'Ziel-Datei füllen
        
            FuelleZielDateiMDE20
            
        Case Is = 2     'Liste 4 leeren
            List4.Clear
            
        Case Is = 3     'Schließe Dialog
            Unload frmWKL93
            
        Case Is = 4     'Auswahlliste wählen
            LeseAuswahlListeMDE20
            
        Case Is = 5     'Auswahlliste speichern
            SpeichereAuswahlListeMDE20
            
        Case Is = 6     'senden
            SendeDaten2MDEGeraetMDE22
            
    End Select
    Screen.MousePointer = vbDefault
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SSCommand1_Click"
        Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub SendeDaten2MDEGeraetMDE22()
    On Error GoTo LOKAL_ERROR
    
    Dim lreti As Integer
    Dim sPfad As String
    Dim sPfad1 As String
    
    sPfad1 = App.Path
    If Right$(sPfad1, 1) <> "\" Then
        sPfad1 = sPfad1 & "\"
    End If
    
    sPfad1 = ShortPath(sPfad1)
    sPfad = ShortPath(App.Path)
    
    lreti = Shell(sPfad & "\FORCOM com" & CStr(gbYtescanPcom) & " +" & sPfad1 & "to_mde.dat s5", vbMinimizedFocus)
   
    

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SendeDaten2MDEGeraetMDE22"
        Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub

Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = vbGreen
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    If Asc(cZeichen) <> vbNull Then
        KeyAscii = Asc(cZeichen)
    Else
        Exit Sub
    End If
    
    cValid = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    If KeyAscii <> 0 And KeyAscii <> 8 Then
        If Len(Text2.Text) = Text2.MaxLength - 1 Then
            SSCommand1(5).SetFocus
        End If
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub Text2_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil BETANKEN ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


