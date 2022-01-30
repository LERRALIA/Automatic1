VERSION 5.00
Begin VB.Form frmWKL74 
   Caption         =   "Kunden Verkauf"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL74.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
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
         TabIndex        =   10
         Top             =   1080
         Width           =   11415
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   6120
         Width           =   9375
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
            Left            =   7080
            TabIndex        =   8
            Tag             =   "Filiale , adate desc"
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
            Left            =   5520
            TabIndex        =   7
            Tag             =   "menge desc"
            Top             =   360
            Width           =   1455
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
            Left            =   3960
            TabIndex        =   6
            Tag             =   "Filiale"
            Top             =   360
            Width           =   1455
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
            Left            =   2160
            TabIndex        =   5
            Tag             =   "Bednr"
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
            Left            =   360
            TabIndex        =   4
            Tag             =   "adate desc"
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
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
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Drucken"
         BeginProperty Font 
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
         Left            =   9600
         TabIndex        =   2
         Top             =   6360
         Width           =   2055
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
         TabIndex        =   11
         Top             =   840
         Width           =   11415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "Artikelanzahl"
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
         Index           =   9
         Left            =   7200
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   6735
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zurück"
      BeginProperty Font 
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
      Top             =   7800
      Width           =   9135
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kunden Verkauf"
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
      TabIndex        =   14
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
Attribute VB_Name = "frmWKL74"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOrder As String

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
        Unload frmWKL74
    Case 1
    
        drucken sOrder
    
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
Private Sub drucken(sOrder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "KUNDD1", gdBase
    CreateTable "KUNDD1", gdBase
    
    sSQL = "Insert into KUNDD1 select "
    sSQL = sSQL & " TEL "
    sSQL = sSQL & ", FAXNR "
    sSQL = sSQL & ", EMAIL "
    sSQL = sSQL & ", MOBILTEL "
    sSQL = sSQL & ", VORNAME "
    sSQL = sSQL & ", KUNDNR "
    sSQL = sSQL & ", NAME "
    sSQL = sSQL & ", STRASSE "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", STADT as ORT "
    sSQL = sSQL & ", TITEL "
    sSQL = sSQL & ", FIRMA "
    sSQL = sSQL & " from KUNDEN where KUNDNR = " & gckundnr
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
   
    loeschNEW "KUNDDRUCK", gdBase
    CreateTable "KUNDDRUCK", gdBase
    
    sSQL = "Insert into KUNDDRUCK select "
    sSQL = sSQL & " ARTNR  "
    sSQL = sSQL & ", BEZEICH  "
    sSQL = sSQL & ", MENGE  "
    sSQL = sSQL & ", PREIS  "
    sSQL = sSQL & ", ADATE  "
    sSQL = sSQL & ", KUNDNR  "
    sSQL = sSQL & ", FILIALE  "
    sSQL = sSQL & ", BEDNR as BEDIENER  "
    
    sSQL = sSQL & " from KUNDAZE "
    sSQL = sSQL & " order by " & sOrder
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
   

    reportbildschirm "", "aWKL74"

    loeschNEW "KUNDD1", gdBase
    loeschNEW "KUNDDRUCK", gdBase


    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren

    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    If gckundnr = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    If IsNumeric(gckundnr) = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    AllesinKUNDAZE gckundnr
    
    List1.AddItem "Datum     Menge   Artnr  Artikelbezeichnung                 Fil Preis    Bediener"
    
    sOrder = "adate desc"
    ZeigArtHistInList "VerkaufKU", List3, gckundnr, sOrder
    
    anzeige "normal", gckundnr, lblAnzeige
'    gckundnr = ""
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
   
End Sub
Private Sub AllesinKUNDAZE(cKund As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "KUNDAZE", gdBase
    CreateTable "KUNDAZE", gdBase
    
    sSQL = "Insert into KUNDAZE select artnr,menge,adate,Filiale,preis,bediener  as bednr from kassjour "
    sSQL = sSQL & " where kundnr = " & cKund
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KUNDAZE select artnr,menge,adate,Filiale,preis,bednr from kundkass "
    sSQL = sSQL & " where kundnr = " & cKund
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNDAZE inner join artikel on kundaze.artnr = artikel.artnr"
    sSQL = sSQL & " set KUNDAZE.Bezeich = artikel.bezeich "
    sSQL = sSQL & " where  artikel.artnr <> 666666"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNDAZE "
    sSQL = sSQL & " set KUNDAZE.Bezeich = 'Gutschein' "
    sSQL = sSQL & " where  KUNDAZE.artnr = 666666"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AllesinKUNDAZE"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten. "
    
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

    sOrder = Option4(Index).Tag
    ZeigArtHistInList "VerkaufKU", List3, gckundnr, sOrder
'    gckundnr = ""

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

