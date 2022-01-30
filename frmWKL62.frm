VERSION 5.00
Begin VB.Form frmWKL62 
   Caption         =   "Artikel Lagerumschlag"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL62.frx":0000
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
      TabIndex        =   2
      Top             =   840
      Width           =   11775
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4155
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   8775
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   8775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "(Umschlagshäufigkeit)"
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
         Left            =   7320
         TabIndex        =   18
         Top             =   5280
         Width           =   4215
      End
      Begin VB.Label Label1 
         Caption         =   "Bestand"
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
         Index           =   3
         Left            =   9120
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Artikelanzahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   9120
         TabIndex        =   16
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Artikelanzahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9120
         TabIndex        =   14
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Lagerreichweite"
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
         Index           =   0
         Left            =   9120
         TabIndex        =   13
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   6240
         TabIndex        =   12
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "Lagerumschlag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   7320
         TabIndex        =   11
         Top             =   4800
         Width           =   4215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   360
         X2              =   5880
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label Label1 
         Caption         =   "Lagerumschlagsdauer"
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
         Index           =   12
         Left            =   9120
         TabIndex        =   10
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Artikelanzahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   9120
         TabIndex        =   9
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         Caption         =   "Artikelanzahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3960
         TabIndex        =   8
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         Caption         =   "Artikelanzahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3960
         TabIndex        =   7
         Top             =   6360
         Width           =   1935
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         Caption         =   "absolute Verkaufsmenge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   600
         TabIndex        =   6
         Top             =   5040
         Width           =   2775
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Rechts
         Caption         =   "durchschnittlicher Bestand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   6000
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         Caption         =   "Artikelanzahl"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   9
         Left            =   7320
         TabIndex        =   4
         Top             =   5760
         Width           =   4215
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
      Left            =   9480
      TabIndex        =   0
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikel Lagerumschlag"
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
Attribute VB_Name = "frmWKL62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKL62
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikel Lagerumschlag ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Positionieren
    
    Modul6.Skalieren Me, True, True:
    Modul6.Schrift Me: Modul6.Log Me
    Modul6.alternativFarbform Me, lblUeberschrift
    
    Dim ctemp As String
    Dim dlug As Double
    Dim dlugd As Double
    
    Screen.MousePointer = 11
    
    dlug = HoleLagerumschlag1(gsARTNR)
    Label1(9).Caption = Format$(dlug, "###0.00")
    Label1(9).Refresh
    
    List1.AddItem "Monat   Jahr         " & Chr(216) & "Bestände  Verkäufe"
    ZeigArtHistInList "UMSCHLAG", List3, gsARTNR, ""
'    gsARTNR = ""

    Label1(6).Caption = Format$(ermSchBest, "###0.0000")
    Label1(6).Refresh
    Label1(7).Caption = Format$(ermSchVerk, "#####0")
    Label1(7).Refresh

    dlugd = 0
    If dlug > 0 Then
        dlugd = 360 / dlug
    End If
    
    Label1(11).Caption = Val(dlugd) & " Tage"
    Label1(11).Refresh
    
    Label1(1).Caption = Val(wievieleTage(gsARTNR)) & " Tage"
    Label1(1).Refresh
    
    Label1(2).Caption = ermBestand(gsARTNR)
    Label1(2).Refresh
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikel Lagerumschlag ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
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
    Fehler.gsFehlertext = "Im Programmteil Artikel Lagerumschlag ist ein Fehler aufgetreten."
    
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
