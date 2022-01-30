VERSION 5.00
Begin VB.Form frmWKL111 
   Caption         =   "Notizen für Artikel"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame2"
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   11655
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6240
         Index           =   2
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   9255
      End
      Begin sevCommand3.Command Command2 
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
         Index           =   11
         Left            =   9480
         TabIndex        =   5
         Top             =   6360
         Width           =   2055
      End
      Begin sevCommand3.Command Command2 
         VBButton        =   1
         ButtonStyle     =   2
         Caption         =   "Speichern"
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
         Index           =   12
         Left            =   9480
         TabIndex        =   4
         Top             =   5760
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Notizen für Artikel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   6855
      End
      Begin VB.Label Label21 
         BackColor       =   &H00008080&
         BackStyle       =   0  'Transparent
         Caption         =   "Notizen für Artikel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   2295
      End
   End
   Begin sevCommand3.Command Command1 
      VBButton        =   1
      ButtonStyle     =   2
      Caption         =   "Schließen"
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
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lbl1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9375
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Notizen für Artikel"
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
      TabIndex        =   1
      Top             =   0
      Width           =   9015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmWKL111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 3
            Unload frmWKL111
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "In Artikel Notizen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    Select Case Index
        Case 11 'Hilfe drucken
            SAVEARTIKEL gsArtNot, Text2(2), Label21(0).Caption
            PrintARTIKEL gsArtNot, Text2(2)
        Case 12 'Hilfe speichern
            SAVEARTIKEL gsArtNot, Text2(2), Label21(0).Caption
            HOLARTIKEL gsArtNot, Text2(2)
        Case 13
            Screen.MousePointer = 0
            Frame5.Visible = False
    End Select
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "In Artikel Notizen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    alternativFarbform Me, lblUeberschrift

    LogtoStart Me
    
    HOLARTIKEL gsArtNot, Text2(2)
    Label21(1).Caption = gsArtNot
    Label21(0).Caption = gsArtNotBez
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "In Artikel Notizen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


