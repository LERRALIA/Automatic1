VERSION 5.00
Begin VB.Form frmWKL78 
   Caption         =   "Artikel Bestandsver�nderung"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL78.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Frame5"
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   11775
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
         Left            =   9480
         TabIndex        =   14
         Top             =   6240
         Width           =   2055
      End
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
         Height          =   4620
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   11535
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
         TabIndex        =   13
         Top             =   240
         Width           =   11535
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   5400
         Width           =   9255
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
            Left            =   5280
            TabIndex        =   8
            Tag             =   "Filiale , LASTDATE desc , lasttime desc"
            Top             =   840
            Value           =   -1  'True
            Width           =   2775
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
            Left            =   5280
            TabIndex        =   7
            Tag             =   "Umenge desc"
            Top             =   480
            Width           =   2415
         End
         Begin VB.OptionButton Option4 
            Caption         =   "�nder. Grund"
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
            Index           =   6
            Left            =   5280
            TabIndex        =   6
            Tag             =   "aenart"
            Top             =   120
            Width           =   3255
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
            Left            =   2040
            TabIndex        =   5
            Tag             =   "Filiale"
            Top             =   840
            Width           =   2655
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
            Left            =   2040
            TabIndex        =   4
            Tag             =   "Bediener"
            Top             =   480
            Width           =   2895
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
            Left            =   2040
            TabIndex        =   3
            Tag             =   "LASTDATE desc ,  lasttime desc"
            Top             =   120
            Width           =   2895
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
            Left            =   120
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
      End
   End
   Begin sevCommand3.Command Command3 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Zur�ck"
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
      Left            =   9480
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
      Left            =   240
      TabIndex        =   12
      Top             =   7800
      Width           =   9135
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikel Bestandsver�nderung"
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
      TabIndex        =   11
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
Attribute VB_Name = "frmWKL78"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sOrder As String
Private Sub Positionieren()
On Error GoTo LOKAL_ERROR
    
    
    With Frame5
        .Height = 6735
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
    Fehler.gsFehlertext = "Im Programmteil Artikel Bestandsver�nderung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            Unload frmWKL78
        Case 1
            drucken sOrder
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Bestandsver�nderung ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub drucken(sOrder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "BESTPDRU" & srechnertab, gdApp
    
    loeschNEW "BESTPDRU", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "BESTPDRU" & srechnertab
    
    sSQL = "select * into BESTPDRU from BESTPDRU" & srechnertab
    gdApp.Execute sSQL, dbFailOnError
    
    reportbildschirmApp "", "aWKL78"

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucken"
    Fehler.gsFehlertext = "Im Programmteil Artikel Bestandsver�nderung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren

    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    sOrder = " LASTDATE desc , lasttime desc"
    List1.AddItem "Datum     Uhrzeit     Fil   ALT  BEW.   NEU  �nderungsart         �nderungsgrund       Bed"
    If SuchedatenAHB(gsARTNR) Then
        ZeigArtHistInList "BESTPDRU", List3, gsARTNR, sOrder
        anzeige "normal", gsARTNR, lblanzeige
    Else
        anzeige "normal", "Keine Daten zu diesem Artikel: " & gsARTNR & " gefunden", lblanzeige
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikel Bestandsver�nderung ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "BESTPDRU" & srechnertab, gdApp
    loeschNEW "BESTPDRU" & srechnertab, gdBase
    
    loeschNEW "BESTPDRU", gdBase
    loeschNEW "BESTPDRU", gdApp
    
    gsARTNR = ""
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
    ZeigArtHistInList "BESTPDRU", List3, gsARTNR, sOrder
'    gsARTNR = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option4_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Bestandsver�nderung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
