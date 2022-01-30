VERSION 5.00
Begin VB.Form frmWKL147 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Kundenbestellungen"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Löschen"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5280
      TabIndex        =   18
      Top             =   7800
      Width           =   2055
   End
   Begin sevCommand3.Command Command3 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Ändern"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   17
      Top             =   7800
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   120
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   7800
      Width           =   2895
   End
   Begin sevCommand3.Command Command3 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Drucken"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7440
      TabIndex        =   14
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   11775
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   11415
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   6120
         Width           =   11295
         Begin VB.OptionButton Option4 
            Caption         =   "Kunde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   9120
            TabIndex        =   13
            Tag             =   "Kundnr"
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Filiale Datum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   6480
            TabIndex        =   7
            Tag             =   "Filiale , bestelltam asc"
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Menge"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   4920
            TabIndex        =   6
            Tag             =   "bestelltmenge desc"
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Filiale"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   3480
            TabIndex        =   5
            Tag             =   "Filiale"
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Bediener"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   4
            Tag             =   "Bednu"
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Datum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   3
            Tag             =   "bestelltam asc"
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Sortierung nach"
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
            Index           =   7
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.ListBox List1 
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
         TabIndex        =   10
         Top             =   840
         Width           =   11415
      End
      Begin VB.Label lblanzeige 
         BackColor       =   &H00C0C000&
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
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label15 
         Caption         =   "Verteilte Artikel, die zur Übertragung bereitstehen"
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
         Left            =   5520
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   5295
      End
   End
   Begin sevCommand3.Command Command3 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Zurück"
      BeginProperty Font  {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kundenbestellungen"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   9495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL147"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    
    With Frame5
        .Height = 6855
        .Left = 0
        .Top = 840
        .Width = 11775
        .BorderStyle = 0
        
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Positionieren"
    Fehler.gsFehlertext = "Im Programmteil verteilte Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 0
        Unload frmWKL147
    Case 1
        KBmLINR "INBESTELLUNG", "", gckundnr
    Case 2
        loescheausKUNDBEST
    Case 3
        UpdateKUNDBEST
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub loescheausKUNDBEST()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz         As String
    Dim cArtNr          As String
    Dim cBestelltam     As String
    Dim cBestelltum     As String
    Dim cSQL            As String
    
    If List3.ListIndex < 0 Then
        MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List3.SetFocus
        Exit Sub
    End If
    
    cLBSatz = List3.list(List3.ListIndex)
    cBestelltam = Left$(cLBSatz, 8)
    cBestelltum = Mid$(cLBSatz, 11, 8)
    cArtNr = Mid$(cLBSatz, 27, 6)
      
    
    cSQL = "Update KUNDBEST set sendok = true  " 'where KUNDNR = " & gcKundnr
    cSQL = cSQL & " , statusartikel  = 'Storno'"
    cSQL = cSQL & " where ARTNR = " & cArtNr
    cSQL = cSQL & " and BESTELLTAM = " & CLng(DateValue(cBestelltam))
    cSQL = cSQL & " and BESTELLTUM = '" & cBestelltum & "'"
    
    gdBase.Execute cSQL, dbFailOnError
    ZeigArtHistInList "KUB1", List3, gckundnr, "bestelltam asc"
     
    anzeige "normal", gckundnr, lblanzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "loescheausKUNDBEST"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub UpdateKUNDBEST()
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz         As String
    Dim cArtNr          As String
    Dim cBestelltam     As String
    Dim cBestelltum     As String
    Dim cSQL            As String
    Dim cSTATUS         As String
    
    
    If List3.ListIndex < 0 Then
        MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Winkiss Hinweis:"
        List3.SetFocus
        Exit Sub
    End If
    
    cLBSatz = List3.list(List3.ListIndex)
    cBestelltam = Left$(cLBSatz, 8)
    cBestelltum = Mid$(cLBSatz, 11, 8)
    cArtNr = Mid$(cLBSatz, 27, 6)
    
    Select Case Combo1.Text
        Case "noch nicht bestellt"
            cSTATUS = "INBESTELLUNG"
        Case "ist bestellt"
            cSTATUS = "BESTELLT"
        Case "geliefert"
            cSTATUS = "GELIEFERT"
        Case "nicht geliefert"
            cSTATUS = "NICHTGELIEFERT"
        Case Else
            MsgBox "Bitte einen Eintrag auswählen!", vbInformation, "Winkiss Hinweis:"
            Combo1.SetFocus
            Exit Sub
        
    End Select
    
        
    cSQL = "Update KUNDBEST set STATUSARTIKEL = '" & cSTATUS & "'"
'    cSQL = cSQL & " Where KUNDNR = " & gcKundnr
    cSQL = cSQL & " Where ARTNR = " & cArtNr
    cSQL = cSQL & " and BESTELLTAM = " & CLng(DateValue(cBestelltam))
    cSQL = cSQL & " and BESTELLTUM = '" & cBestelltum & "'"
    gdBase.Execute cSQL, dbFailOnError
        
    ZeigArtHistInList "KUB1", List3, gckundnr, "bestelltam asc"
'    ZeigArtHistInList "KUBE", List3, gcKundnr, "StatusARTIKEL"
    anzeige "normal", gckundnr, lblanzeige
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateKUNDBEST"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren

    alternativFarbform Me, lblUeberschrift
'    Skalieren Me, True, True: Schrift Me
    LogtoStart Me
    
    List1.AddItem "Datum     Uhrzeit Menge   Artnr  Artikelbezeichnung                 Fil Preis    KundNr  Bed."
    
    ZeigArtHistInList "KUB1", List3, gckundnr, "bestelltam asc"
    anzeige "normal", gckundnr, lblanzeige
    
    
    fuellecombo1

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
   
End Sub
Private Sub fuellecombo1()
    On Error GoTo LOKAL_ERROR
    
    Combo1.Clear
    Combo1.AddItem "noch nicht bestellt"
    Combo1.AddItem "ist bestellt"
    Combo1.AddItem "geliefert"
    Combo1.AddItem "nicht geliefert"
    Combo1.AddItem "bitte auswählen"
    
    Combo1.Text = "bitte auswählen"
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo1"
    Fehler.gsFehlertext = "Im Programmteil Kundenbestellungen ist ein Fehler aufgetreten. "
    
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

Private Sub Option4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    ZeigArtHistInList "KUB1", List3, gckundnr, Option4(Index).Tag

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


