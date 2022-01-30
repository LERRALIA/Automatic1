VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmWKLat 
   BackColor       =   &H00C0C000&
   Caption         =   "Bediener Statistik"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   630
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
   ForeColor       =   &H00404000&
   Icon            =   "frmWKLat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   6240
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   250
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   5355
      TabIndex        =   37
      Top             =   550
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.ComboBox cboBed 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   35
      Text            =   "alle"
      Top             =   480
      Width           =   3495
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   7920
      Visible         =   0   'False
      Width           =   3255
      Begin VB.OptionButton optK 
         BackColor       =   &H00C0C000&
         Caption         =   "Kunden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   2160
         TabIndex        =   32
         ToolTipText     =   "sortiert nach Kunden, verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag"
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optA 
         BackColor       =   &H00C0C000&
         Caption         =   "AGN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1440
         TabIndex        =   31
         ToolTipText     =   "sortiert nach AGN, verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag"
         Top             =   0
         Width           =   735
      End
      Begin VB.OptionButton optL 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferanten"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "sortiert nach Lieferanten, verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag"
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   735
      Left            =   360
      TabIndex        =   25
      Top             =   7200
      Width           =   6255
      Begin VB.OptionButton optD 
         BackColor       =   &H00C0C000&
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   3120
         TabIndex        =   33
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton optq 
         BackColor       =   &H00C0C000&
         Caption         =   "einfache Übersicht"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag"
         Top             =   0
         Width           =   2535
      End
      Begin VB.OptionButton optqp 
         BackColor       =   &H00C0C000&
         Caption         =   "erweiterte Übersicht"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag, %Anteil Umsatz, Kundenschnitt in €, Anzahl Kunden"
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton optz 
         BackColor       =   &H00C0C000&
         Caption         =   "Zusammenfassung nach"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   3120
         TabIndex        =   26
         ToolTipText     =   "sortiert nach Lieferanten, verkaufte Artikel, Umsatz(VK), Umsatz(EK), Ertrag"
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      Picture         =   "frmWKLat.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   17
      ToolTipText     =   "Löschen Ihrer Eingaben"
      Top             =   7440
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   735
      Left            =   480
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1296
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Schließen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   7920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Drucken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   7440
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar pbrZeit 
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   6960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSACAL.Calendar caldate 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
      _Version        =   524288
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2002
      Month           =   4
      Day             =   30
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   2
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   -1  'True
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Auswahl einschränken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   480
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   7935
      Begin VB.ComboBox cboLief 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Text            =   "alle Lieferanten"
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox cboAgn 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   24
         Text            =   "alle AGN´s"
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox cbodat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   20
         Text            =   "Zeitraum auswählen"
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cboKunde 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   18
         Text            =   "alle Kunden"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cboLin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6240
         TabIndex        =   16
         Text            =   "alle Linien"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "sortiert nach"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   480
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   6855
      Begin VB.ComboBox cboSort4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown-Liste
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboSort3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboSort2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown-Liste
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboSort1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmWKLat.frx":0804
         Left            =   120
         List            =   "frmWKLat.frx":0814
         Style           =   2  'Dropdown-Liste
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdQuick 
      BackColor       =   &H00C0C000&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   10850
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Grafisch
      TabIndex        =   19
      ToolTipText     =   "Starten Sie hier die Anzeige"
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1695
      Left            =   8520
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton cmdListDel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   1080
         Width           =   375
      End
      Begin VB.ListBox lstLinA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lblAnzeige 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   34
      Top             =   6960
      Width           =   10815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      X1              =   480
      X2              =   11280
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Bediener - Statistik"
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
      Left            =   480
      TabIndex        =   9
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "bis:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8520
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "von:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuAusw 
      Caption         =   "vorgefertigte Auswertungen"
      Index           =   0
      Begin VB.Menu mnuBed 
         Caption         =   "...Verkauf pro Kunde"
         Index           =   0
      End
      Begin VB.Menu mnuBed 
         Caption         =   "...Entwicklung Verkauf pro Kunde"
         Index           =   1
      End
      Begin VB.Menu mnuBed 
         Caption         =   "...Umsatz pro Kunde"
         Index           =   2
      End
      Begin VB.Menu mnuBed 
         Caption         =   "...Ertrag pro Kunde"
         Index           =   3
      End
      Begin VB.Menu mnuBed 
         Caption         =   "...Provisionen"
         Index           =   4
      End
      Begin VB.Menu mnuBed 
         Caption         =   "...Provisionen rabattierfähiger Artikel"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmWKLat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Public iPruef As Integer
Public bLinie As Boolean
Public gbNoData As Boolean
Public bFlexQ As Boolean
Public bsortOrAusw As Boolean
Public bQuick As Boolean
Public bQuickPlus As Boolean
Public bZ As Boolean
Public bDetail As Boolean

Public iLieferant As Integer
Public iLinie As Integer
Public iAGN As Integer
Public iKunde As Integer
Public iDatum As Integer
Private Sub caldate_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim dteVon As Date
    Dim dteBis As Date
    
    lblAnzeige.Caption = ""

    If iPruef = 1 Then
        Text1(0).Text = caldate.Value
        dteVon = DateValue(Text1(0).Text)
        If dteVon > Date Then
            lblAnzeige.Caption = "Das heutige Datum ist überschritten worden."
            lblAnzeige.Refresh
            Text1(0).Text = ""
        End If
        Text1(0).SetFocus
    ElseIf iPruef = 2 Then
        If Text1(0).Text <> "" Then
            Text1(1).Text = caldate.Value
            
            dteVon = DateValue(Text1(0).Text)
            dteBis = DateValue(Text1(1).Text)
            
            If dteBis > Date Then
                lblAnzeige.Caption = "Das heutige Datum ist überschritten worden."
                lblAnzeige.Refresh
                Text1(1).Text = ""
            ElseIf dteBis < dteVon Then
                lblAnzeige.Caption = "Das Datum für das Ende des gewählten Zeitraums ist kleiner als das Anfangsdatum."
                lblAnzeige.Refresh
                Text1(1).Text = ""
            End If
            Text1(1).SetFocus
        Else
            lblAnzeige.Caption = "Geben Sie bitte erst ein Anfangsdatum ein!"
            lblAnzeige.Refresh
            Text1(0).SetFocus
        End If
                
    End If
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "caldate_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub
Private Sub frame2create1()
    On Error GoTo LOKAL_ERROR

    ' Lieferanten
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 3975
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
''    cboLief.Height = 315
    cboLief.Left = 120
    cboLief.Top = 240
    cboLief.Width = 3735
    
    cboKunde.Visible = False
    cboAgn.Visible = False
    cboLin.Visible = False
    cbodat.Visible = False
    
    iLieferant = 2
    iLinie = 3
    iKunde = 4
    iAGN = 5
    iDatum = 6
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create1"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create2()
    On Error GoTo LOKAL_ERROR

    ' Kunden
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 4095
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboKunde.Text = "alle Kunden"
    cboKunde.Visible = True
'    cboKunde.Height = 315
    cboKunde.Left = 120
    cboKunde.Top = 240
    cboKunde.Width = 3855
    
    cboLin.Visible = False
    cboLief.Visible = False
    cboAgn.Visible = False
    cbodat.Visible = False
    
    iLieferant = 3
    iLinie = 4
    iKunde = 2
    iAGN = 5
    iDatum = 6
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create2"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create3()
    On Error GoTo LOKAL_ERROR
    
    ' AGN
    
    Frame3.Caption = "AGN-Auswahl"
    Frame3.Height = 1695
    Frame3.Left = 4560
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 3975
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboAgn.Text = "alle AGN´s"
    cboAgn.Visible = True
'    cboAgn.Height = 315
    cboAgn.Left = 120
    cboAgn.Top = 240
    cboAgn.Width = 3735
    
    cboLief.Visible = False
    cboKunde.Visible = False
    cboLin.Visible = False
    cbodat.Visible = False
    
    iLieferant = 3
    iLinie = 4
    iKunde = 5
    iAGN = 2
    iDatum = 6
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create3"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create3a()
    On Error GoTo LOKAL_ERROR

    ' Datum
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 3975
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cbodat.Text = "Zeitraum auswählen"
    cbodat.Visible = True
'    cbodat.Height = 315
    cbodat.Left = 120
    cbodat.Top = 240
    cbodat.Width = 3735
    
    cboLief.Visible = False
    cboKunde.Visible = False
    cboLin.Visible = False
    cboKunde.Visible = False
    
    iLieferant = 3
    iLinie = 4
    iKunde = 5
    iAGN = 6
    iDatum = 2
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create3a"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create4()
    On Error GoTo LOKAL_ERROR

    ' Lieferanten und Linie
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 6855
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
    
    cboLin.Clear
    cboLin.Text = "alle Linien"
    cboLin.Visible = True
    
    
'    cboLief.Height = 315
    cboLief.Left = 120
    cboLief.Top = 240
    cboLief.Width = 3735
    
'    cboLin.Height = 315
    cboLin.Left = 3960
    cboLin.Top = 240
    cboLin.Width = 2775
    
    cboKunde.Visible = False
    cboAgn.Visible = False
    cbodat.Visible = False
    bLinie = True
    
    iLieferant = 2
    iLinie = 3
    iKunde = 4
    iAGN = 5
    iDatum = 6
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create4"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create5()
    On Error GoTo LOKAL_ERROR

    ' Lieferanten und AGN
    
    Frame3.Caption = "AGN-Auswahl"
    Frame3.Height = 1695
    Frame3.Left = 8400
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7815
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
    
    cboAgn.Text = "alle AGN´s"
    cboAgn.Visible = True
    
'    cboLief.Height = 315
    cboLief.Left = 120
    cboLief.Top = 240
    cboLief.Width = 3735
    
'    cboAgn.Height = 315
    cboAgn.Left = 3960
    cboAgn.Top = 240
    cboAgn.Width = 3735
    
    cboKunde.Visible = False
    cboLin.Visible = False
    cbodat.Visible = False
    
    iLieferant = 2
    iLinie = 4
    iKunde = 5
    iAGN = 3
    iDatum = 6
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create5"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create5b()
    On Error GoTo LOKAL_ERROR

    ' Datum und AGN
    Frame3.Caption = "AGN-Auswahl"
    Frame3.Height = 1695
    Frame3.Left = 8400
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7815
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cbodat.Text = "Zeitraum auswählen"
    cbodat.Visible = True
    
    cboAgn.Text = "alle AGN´s"
    cboAgn.Visible = True
    
'    cbodat.Height = 315
    cbodat.Left = 120
    cbodat.Top = 240
    cbodat.Width = 3735
    
'    cboAgn.Height = 315
    cboAgn.Left = 3960
    cboAgn.Top = 240
    cboAgn.Width = 3735
    
    cboKunde.Visible = False
    cboLin.Visible = False
    cboLief.Visible = False
    
    iLieferant = 4
    iLinie = 5
    iKunde = 6
    iAGN = 3
    iDatum = 2
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create5b"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create5a()
    On Error GoTo LOKAL_ERROR

    ' Lieferanten und Datum
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7815
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
    
    cbodat.Text = "Zeitraum auswählen"
    cbodat.Visible = True
    
'    cboLief.Height = 315
    cboLief.Left = 120
    cboLief.Top = 240
    cboLief.Width = 3735
    
'    cbodat.Height = 315
    cbodat.Left = 3960
    cbodat.Top = 240
    cbodat.Width = 3735
    
    cboKunde.Visible = False
    cboLin.Visible = False
    cboAgn.Visible = False
    
    iLieferant = 2
    iLinie = 5
    iKunde = 6
    iAGN = 4
    iDatum = 3
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create5a"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create6()
    On Error GoTo LOKAL_ERROR

    ' Lieferanten und Kunden
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7935
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
    
    cboKunde.Text = "alle Kunden"
    cboKunde.Visible = True
    
'    cboLief.Height = 315
    cboLief.Left = 120
    cboLief.Top = 240
    cboLief.Width = 3735
    
'    cboKunde.Height = 315
    cboKunde.Left = 3960
    cboKunde.Top = 240
    cboKunde.Width = 3855
    
    cboLin.Visible = False
    cboAgn.Visible = False
    cbodat.Visible = False
    
    iLieferant = 2
    iLinie = 4
    iKunde = 3
    iAGN = 5
    iDatum = 6
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create6"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create7()
    On Error GoTo LOKAL_ERROR

    ' AGN und Lieferanten
    Frame3.Caption = "AGN-Auswahl"
    Frame3.Height = 1695
    Frame3.Left = 8400
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7815
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
    
'    cboLief.Height = 315
    cboLief.Left = 3960
    cboLief.Top = 240
    cboLief.Width = 3735
    
   cboAgn.Text = "alle AGN´s"
    cboAgn.Visible = True
'    cboAgn.Height = 315
    cboAgn.Left = 120
    cboAgn.Top = 240
    cboAgn.Width = 3735
    
    cboLin.Visible = False
    cboKunde.Visible = False
    cbodat.Visible = False
    
    iLieferant = 3
    iLinie = 4
    iKunde = 5
    iAGN = 2
    iDatum = 6
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "caldate_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create7b()
    On Error GoTo LOKAL_ERROR

    ' Datum und Lieferanten
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7815
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
    
'    cboLief.Height = 315
    cboLief.Left = 3960
    cboLief.Top = 240
    cboLief.Width = 3735
    
    cbodat.Text = "Zeitraum auswählen"
    cbodat.Visible = True
'    cbodat.Height = 315
    cbodat.Left = 120
    cbodat.Top = 240
    cbodat.Width = 3735
    
    cboLin.Visible = False
    cboKunde.Visible = False
    cboAgn.Visible = False
    
    iLieferant = 3
    iLinie = 4
    iKunde = 5
    iAGN = 6
    iDatum = 2
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create7b"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create7a()
    On Error GoTo LOKAL_ERROR

    ' AGN und Datum
    Frame3.Caption = "AGN-Auswahl"
    Frame3.Height = 1695
    Frame3.Left = 8400
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7815
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cbodat.Text = "Zeitraum auswählen"
    cbodat.Visible = True
    
'    cbodat.Height = 315
    cbodat.Left = 3960
    cbodat.Top = 240
    cbodat.Width = 3735
    
    cboAgn.Text = "alle AGN´s"
    cboAgn.Visible = True
'    cboAgn.Height = 315
    cboAgn.Left = 120
    cboAgn.Top = 240
    cboAgn.Width = 3735
    
    cboLin.Visible = False
    cboKunde.Visible = False
    cboLief.Visible = False
    
    
    iLieferant = 4
    iLinie = 5
    iKunde = 6
    iAGN = 2
    iDatum = 3
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create7a"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create8()
    On Error GoTo LOKAL_ERROR

    ' AGN und Kunden
    Frame3.Caption = "AGN-Auswahl"
    Frame3.Height = 1695
    Frame3.Left = 8520
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7935
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboAgn.Text = "alle AGN´s"
    cboAgn.Visible = True
    
'    cboAgn.Height = 315
    cboAgn.Left = 120
    cboAgn.Top = 240
    cboAgn.Width = 3735
    
    cboKunde.Text = "alle Kunden"
    cboKunde.Visible = True
    
'    cboKunde.Height = 315
    cboKunde.Left = 3960
    cboKunde.Top = 240
    cboKunde.Width = 3855
    
    cboLin.Visible = False
    cboLief.Visible = False
    cbodat.Visible = False
    
    iLieferant = 4
    iLinie = 5
    iKunde = 3
    iAGN = 2
    iDatum = 6
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create8"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create8b()
    On Error GoTo LOKAL_ERROR

    ' Datum und Kunden
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7935
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cbodat.Text = "Zeitraum auswählen"
    cbodat.Visible = True
    
'    cbodat.Height = 315
    cbodat.Left = 120
    cbodat.Top = 240
    cbodat.Width = 3735
    
    cboKunde.Text = "alle Kunden"
    cboKunde.Visible = True
    
'    cboKunde.Height = 315
    cboKunde.Left = 3960
    cboKunde.Top = 240
    cboKunde.Width = 3855
    
    cboLin.Visible = False
    cboLief.Visible = False
    cboAgn.Visible = False
    
    iLieferant = 4
    iLinie = 5
    iKunde = 3
    iAGN = 6
    iDatum = 2
    
    
    
Exit Sub
LOKAL_ERROR:
   Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create8b"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create9()
    On Error GoTo LOKAL_ERROR
    
    ' Kunden und AGN
    Frame3.Caption = "AGN-Auswahl"
    Frame3.Height = 1695
    Frame3.Left = 8520
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7935
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboKunde.Text = "alle Kunden"
    cboKunde.Visible = True
    
    cboKunde.Left = 120
    cboKunde.Top = 240
    cboKunde.Width = 3855
    
    cboAgn.Text = "alle AGN´s"
    cboAgn.Visible = True
    
    cboAgn.Left = 4080
    cboAgn.Top = 240
    cboAgn.Width = 3735
    
    cboLin.Visible = False
    cboLief.Visible = False
    cbodat.Visible = False
    
    iLieferant = 4
    iLinie = 5
    iKunde = 2
    iAGN = 3
    iDatum = 6
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create9"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create9a()
    On Error GoTo LOKAL_ERROR

    ' Kunden und Datum
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7935
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboKunde.Text = "alle Kunden"
    cboKunde.Visible = True
    
    cboKunde.Left = 120
    cboKunde.Top = 240
    cboKunde.Width = 3855
    
    cbodat.Text = "Zeitraum auswählen"
    cbodat.Visible = True
    
    cbodat.Left = 4080
    cbodat.Top = 240
    cbodat.Width = 3735
    
    cboLin.Visible = False
    cboLief.Visible = False
    cboAgn.Visible = False
    
    iLieferant = 5
    iLinie = 6
    iKunde = 2
    iAGN = 4
    iDatum = 3
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create9a"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2create10()
    On Error GoTo LOKAL_ERROR

    ' Kunden und Lieferanten
    
    Frame2.Visible = True
    Frame2.Height = 735
    Frame2.Left = 480
    Frame2.Top = 1920
    Frame2.Width = 7935
    
    Modul6.SkalierenFrame Frame2, True, True
    
    cboLief.Text = "alle Lieferanten"
    cboLief.Visible = True
    
    cboKunde.Text = "alle Kunden"
    cboKunde.Visible = True
    

    cboLief.Left = 4080
    cboLief.Top = 240
    cboLief.Width = 3735
    
    cboKunde.Left = 120
    cboKunde.Top = 240
    cboKunde.Width = 3855
    
    cboLin.Visible = False
    cboAgn.Visible = False
    cbodat.Visible = False
    
    iLieferant = 3
    iLinie = 4
    iKunde = 2
    iAGN = 5
    iDatum = 6
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2create10"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub framegroesse1()
    On Error GoTo LOKAL_ERROR

    Frame1.Visible = True
    Frame1.Height = 735
    Frame1.Left = 480
    Frame1.Top = 960
    Frame1.Width = 1815
    
    Modul6.SkalierenFrame Frame1, True, True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "framegroesse1"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub framegroesse2()
    On Error GoTo LOKAL_ERROR

    Frame1.Visible = True
    Frame1.Height = 735
    Frame1.Left = 480
    Frame1.Top = 960
    Frame1.Width = 3495
    
    Modul6.SkalierenFrame Frame1, True, True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "framegroesse2"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub framegroesse3()
    On Error GoTo LOKAL_ERROR

    Frame1.Visible = True
    Frame1.Height = 735
    Frame1.Left = 480
    Frame1.Top = 960
    Frame1.Width = 5175
    
    Modul6.SkalierenFrame Frame1, True, True
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "framegroesse3"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub framegroesse4()
    On Error GoTo LOKAL_ERROR

    Frame1.Visible = True
    Frame1.Height = 735
    Frame1.Left = 480
    Frame1.Top = 960
    Frame1.Width = 6855
    
    Modul6.SkalierenFrame Frame1, True, True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "framegroesse4"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboAgn_Click()
    On Error GoTo LOKAL_ERROR

    Dim sAusw As String
    bsortOrAusw = True
    
    Frame3.Visible = True
    
    
    
    sAusw = Trim(Right(cboAgn.Text, 4))
    lstLinA.AddItem (sAusw)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboAgn_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub









Private Sub cbobed_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    caldate.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbobed_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub cbodat_Click()
    On Error GoTo LOKAL_ERROR

    Dim imon As Integer
    Dim iyear As Integer
    
    Dim cVon As String
    Dim cBis As String
    
    Dim lBis As Long
    
    bsortOrAusw = True
    
    Select Case cbodat.Text
        Case Is = "Zeitraum auswählen"
            Text1(0).Text = DateValue(Now)
            Text1(1).Text = DateValue(Now)
        Case Is = "Heute"
            Text1(0).Text = DateValue(Now)
            Text1(1).Text = DateValue(Now)
        Case Is = "Gestern"
            Text1(0).Text = DateValue(Now) - 1
            Text1(1).Text = DateValue(Now) - 1
        Case Is = "letzten 3 Tage"
            Text1(0).Text = DateValue(Now) - 3
            Text1(1).Text = DateValue(Now)
        Case Is = "letzte Woche"
            Text1(0).Text = DateValue(Now) - 6
            Text1(1).Text = DateValue(Now)
        
        Case Is = "letzter Monat"
            imon = Month(Now)
            iyear = Year(Now)
            If imon = 1 Then
            imon = 12
            iyear = Year(Now) - 1
            Else
            imon = imon - 1
            End If
            If imon = 1 Or imon = 3 Or imon = 5 Or imon = 7 Or imon = 8 Or imon = 10 Or imon = 12 Then
            lBis = 31
            ElseIf imon = 2 Then lBis = 28
            Else: lBis = 30
            End If
            
            cVon = "01." & imon & "." & iyear
            cBis = lBis & "." & imon & "." & iyear
            Text1(0).Text = DateValue(cVon)
            Text1(1).Text = DateValue(cBis)
        Case Is = "dieser Monat"
            imon = Month(Now)
            If imon = 1 Or imon = 3 Or imon = 5 Or imon = 7 Or imon = 8 Or imon = 10 Or imon = 12 Then
            lBis = 31
            ElseIf imon = 2 Then lBis = 28
            Else: lBis = 30
            End If
            cVon = "01." & Month(Now) & "." & Year(Now)
            cBis = lBis & "." & Month(Now) & "." & Year(Now)
            Text1(0).Text = DateValue(cVon)
            Text1(1).Text = DateValue(cBis)
        Case Is = "letzten 30 Tage"
            Text1(0).Text = DateValue(Now) - 30
            Text1(1).Text = DateValue(Now)
        
        Case Else
            'nix tun
        End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbodat_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cbodat_GotFocus()
On Error GoTo LOKAL_ERROR

fülledat

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbodat_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    End Sub

Private Sub cboKunde_Click()
    On Error GoTo LOKAL_ERROR

    bsortOrAusw = True
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboKunde_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub cboKunde_GotFocus()
    On Error GoTo LOKAL_ERROR

    füllecboKunden
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboKunde_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
        
End Sub

Private Sub cboLief_Click()
    On Error GoTo LOKAL_ERROR

    Dim sLieferant As String
    
    bsortOrAusw = True
    
    If bLinie Then
        If cboLief.Text <> "" Then
            sLieferant = cboLief.Text
            füllecboLinie (sLieferant)
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboLief_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub füllecboLinie(sLieferant As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "SELECT LINBEZEICH,lpz FROM LINBEZ INNER JOIN LISRT ON LINBEZ.LINR = LISRT.LINR"
    sSQL = sSQL & " Where LISRT.LIEFBEZ = '" & sLieferant & "' "
    sSQL = sSQL & " order BY LINBEZ.LINBEZEICH "
    
    Set rs = gdBase.OpenRecordset(sSQL)
    
    cboLin.Clear
   
    Do While Not rs.EOF
        cboLin.AddItem rs!LINBEZEICH & Space(80 - Len(rs!LINBEZEICH)) & rs!LPZ
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    cboLin.Text = "alle Linien"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllecboLinie"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub füllecboAgn()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "select distinct agtext,agn from agndbf where agtext is not null order by agtext"
    
    Set rs = gdBase.OpenRecordset(sSQL)
    
    cboAgn.Clear
   
    Do While Not rs.EOF
        cboAgn.AddItem rs!AGTEXT & Space(36 - Len(rs!AGTEXT)) & rs!AGN
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    cboAgn.Text = "alle AGN´s"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllecboAgn"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboLief_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    füllecboLieferanten
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboLief_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cbolief_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    Dim llen As Long
    Dim Res As Long
   
    ' Eingaben nur für Zahlen und Buchstapen prüfen
    If KeyCode >= vbKey0 And KeyCode <= vbKeyZ Then
       
       ' der ByVal Aufruf des Suchstrings ist eine Eigenheit von Windows/VB
       ' bezüglich der Stringpointer. Nur so wird ein gültiger Pointer auf den
       ' String übergeben. Ohne ByVal geht's nicht!!!
       
       With cboLief
          ' Eintrag suchen: Die CB_SELECTSTRING Message ist dafür ungeignet,
          ' da sie den String sofort einträgt und komplett markiert
          
          Res = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal .Text)
    
          If Res >= 0 Then     ' Eintrag gefunden
             'Call SendMessage(.hwnd, WM_SETREDRAW, False, 0&)
             llen = Len(.Text)  ' Laenge des eingegeben Textes
             .ListIndex = Res   ' Listindex auf gefunden Text setzen
             .Text = .list(Res) ' gefunden Text in Textfeld eintragen
             .SelStart = llen   ' Textende markieren, damit es bei neuer Eingabe gleich wieder gelöscht wird
             .SelLength = Len(.Text) - llen
             'Call SendMessage(.hwnd, WM_SETREDRAW, True, 0&)
          End If
       End With
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cbolief_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboLin_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sAusw As String
    bsortOrAusw = True
    
    Frame3.Visible = True
    Frame3.Caption = "Linienauswahl"
    Frame3.Height = 1695
    Frame3.Left = 7440
    Frame3.Top = 960
    Frame3.Width = 1455
    
    Modul6.SkalierenFrame Frame3, True, True
    
    
    
    sAusw = Trim(Right(cboLin.Text, 4))
    lstLinA.AddItem (sAusw)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboLin_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboSort1_Click()
    On Error GoTo LOKAL_ERROR
    
    füllecboAgn
    
    caldate.Visible = False
    Frame3.Visible = False
    lstLinA.Clear
    
    If cboSort1.Text = "Lieferanten" Then
    
        frame2create1
        cboSort2.Visible = True
        cboSort2.Clear
        cboSort2.Refresh
        cboSort2.AddItem "Linie"
        cboSort2.AddItem "AGN"
        cboSort2.AddItem "Kunden"
        cboSort2.AddItem "Datum"
        framegroesse2
        
    ElseIf cboSort1.Text = "AGN" Then
    
    
        frame2create3
        
        cboSort2.Visible = True
        cboSort2.Clear
        cboSort2.Refresh
        cboSort2.AddItem "Lieferanten"
        cboSort2.AddItem "Kunden"
        cboSort2.AddItem "Datum"
        framegroesse2
        
    ElseIf cboSort1.Text = "Kunden" Then
        
        frame2create2
        
        cboSort2.Visible = True
        cboSort2.Clear
        cboSort2.Refresh
        cboSort2.AddItem "AGN"
        cboSort2.AddItem "Lieferanten"
        cboSort2.AddItem "Datum"
        framegroesse2
        
    ElseIf cboSort1.Text = "Datum" Then
        
        frame2create3a
        
        cboSort2.Visible = True
        cboSort2.Clear
        cboSort2.Refresh
        cboSort2.AddItem "AGN"
        cboSort2.AddItem "Lieferanten"
        cboSort2.AddItem "Kunden"
        framegroesse2
        
       
        
    ElseIf cboSort1.Text = "" Then
        cboSort2.Visible = False
        framegroesse1
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboSort1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboSort1_GotFocus()
    On Error GoTo LOKAL_ERROR

    bsortOrAusw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboSort1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboSort2_Click()
    On Error GoTo LOKAL_ERROR
    
        füllecboAgn
        
        caldate.Visible = False
        Frame3.Visible = False
        lstLinA.Clear
        '#####################
        'Lieferanten
        
        'Lieferant,Linie
        If cboSort2.Text = "Linie" Then
            frame2create4
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "AGN"
            cboSort3.AddItem "Kunden"
            cboSort3.AddItem "Datum"
            framegroesse3
        
        'Lieferant,AGN
        ElseIf cboSort2.Text = "AGN" And cboSort1.Text = "Lieferanten" Then
            frame2create5
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Linie"
            cboSort3.AddItem "Kunden"
            cboSort3.AddItem "Datum"
            framegroesse3
        
        'Lieferant,Kunde
        ElseIf cboSort2.Text = "Kunden" And cboSort1.Text = "Lieferanten" Then
            frame2create6
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Linie"
            cboSort3.AddItem "AGN"
            cboSort3.AddItem "Datum"
            framegroesse3
        
        'Lieferant,Datum
        ElseIf cboSort2.Text = "Datum" And cboSort1.Text = "Lieferanten" Then
            frame2create5a
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Linie"
            cboSort3.AddItem "AGN"
            cboSort3.AddItem "Kunden"
            framegroesse3
        
        
        
        '##############################
        'AGN
        
        'AGN,Lieferanten
        ElseIf cboSort2.Text = "Lieferanten" And cboSort1.Text = "AGN" Then
            frame2create7
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Linie"
            cboSort3.AddItem "Kunden"
            cboSort3.AddItem "Datum"
            framegroesse3
        
        'AGN, Kunden
        ElseIf cboSort2.Text = "Kunden" And cboSort1.Text = "AGN" Then
            frame2create8
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Lieferanten"
            cboSort3.AddItem "Datum"
            framegroesse3
        
        'AGN,Datum
        ElseIf cboSort2.Text = "Datum" And cboSort1.Text = "AGN" Then
            frame2create7a
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Lieferanten"
            cboSort3.AddItem "Kunden"
            framegroesse3
        
        '########################
        'Kunden
        
        'Kunden,AGN
        ElseIf cboSort2.Text = "AGN" And cboSort1.Text = "Kunden" Then
            frame2create9
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Lieferanten"
            cboSort3.AddItem "Datum"
            framegroesse3
        
        'Kunden,Lieferanten
        ElseIf cboSort2.Text = "Lieferanten" And cboSort1.Text = "Kunden" Then
            frame2create10
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Linie"
            cboSort3.AddItem "AGN"
            cboSort3.AddItem "Datum"
            framegroesse3
        
        'Kunden,Datum
        ElseIf cboSort2.Text = "Datum" And cboSort1.Text = "Kunden" Then
            frame2create9a
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Lieferanten"
            cboSort3.AddItem "AGN"
            framegroesse3
        
        '####################
        'Datum
        
        'Datum,Lieferanten
        ElseIf cboSort2.Text = "Lieferanten" And cboSort1.Text = "Datum" Then
            frame2create7b
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Linie"
            cboSort3.AddItem "AGN"
            cboSort3.AddItem "Kunden"
            framegroesse3
        
        'Datum,Kunden
        ElseIf cboSort2.Text = "Kunden" And cboSort1.Text = "Datum" Then
            frame2create8b
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Lieferanten"
            cboSort3.AddItem "AGN"
            framegroesse3
            
        'Datum,AGN
        ElseIf cboSort2.Text = "AGN" And cboSort1.Text = "Datum" Then
            frame2create5b
            cboSort3.Visible = True
            cboSort3.Clear
            cboSort3.Refresh
            cboSort3.AddItem "Lieferanten"
            cboSort3.AddItem "Kunden"
            framegroesse3
             
        ElseIf cboSort2.Text = "" Then
            cboSort3.Visible = False
            framegroesse2
            
        End If
        
        Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboSort2_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboSort3_Click()
    On Error GoTo LOKAL_ERROR
    
    caldate.Visible = False
    'Lieferant,Linie,Datum
    If cboSort3.Text = "Datum" And cboSort2.Text = "Linie" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "AGN"
        framegroesse4
        
        iLieferant = 2
        iLinie = 3
        iDatum = 4
        iKunde = 6
        iAGN = 5
        
       
        
    'Lieferant,Linie,Kunden
    ElseIf cboSort3.Text = "Kunden" And cboSort2.Text = "Linie" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 2
        iLinie = 3
        iKunde = 4
        iAGN = 5
        iDatum = 6
    'Lieferant,linie,AGN
    ElseIf cboSort3.Text = "AGN" And cboSort2.Text = "Linie" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 2
        iLinie = 3
        iKunde = 5
        iAGN = 4
        iDatum = 6
        
    'Lieferant,Kunden,Linie
    ElseIf cboSort3.Text = "Linie" And cboSort2.Text = "Kunden" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 2
        iLinie = 4
        iKunde = 3
        iAGN = 5
        iDatum = 6
        
    'Lieferant,Kunden,AGN
    ElseIf cboSort3.Text = "AGN" And cboSort2.Text = "Kunden" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 2
        iLinie = 5
        iKunde = 3
        iAGN = 4
        iDatum = 6
        
    'Lieferant,Kunden,Datum
    ElseIf cboSort3.Text = "Datum" And cboSort2.Text = "Kunden" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 2
        iLinie = 5
        iKunde = 3
        iAGN = 6
        iDatum = 4
        
    
  
    'Lieferant,AGN,Linie
    ElseIf cboSort3.Text = "Linie" And cboSort2.Text = "AGN" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 2
        iLinie = 4
        iKunde = 5
        iAGN = 3
        iDatum = 6
        
    'Lieferant,AGN,Kunde
    ElseIf cboSort3.Text = "Kunden" And cboSort2.Text = "AGN" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 2
        iLinie = 5
        iKunde = 4
        iAGN = 3
        iDatum = 6
        
    'Lieferant,AGN,Datum
    ElseIf cboSort3.Text = "Datum" And cboSort2.Text = "AGN" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 2
        iLinie = 6
        iKunde = 5
        iAGN = 3
        iDatum = 4
    
    'Lieferant,Datum,Linie
    ElseIf cboSort3.Text = "Linie" And cboSort2.Text = "Datum" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "AGN"
        framegroesse4
        
        iLieferant = 2
        iLinie = 4
        iKunde = 5
        iAGN = 6
        iDatum = 3
        
    'Lieferant,Datum,Kunden
    ElseIf cboSort3.Text = "Kunden" And cboSort2.Text = "Datum" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "AGN"
        framegroesse4
        
        iLieferant = 2
        iLinie = 5
        iKunde = 4
        iAGN = 6
        iDatum = 3
        
    'Lieferant,Datum,AGN
    ElseIf cboSort3.Text = "AGN" And cboSort2.Text = "Datum" And cboSort1.Text = "Lieferanten" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 2
        iLinie = 5
        iKunde = 6
        iAGN = 4
        iDatum = 3
    
    'AGN,Lieferant,Linie
    ElseIf cboSort3.Text = "Linie" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "AGN" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 3
        iLinie = 4
        iKunde = 5
        iAGN = 2
        iDatum = 6
        
    'AGN,Lieferant,Kunden
    ElseIf cboSort3.Text = "Kunden" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "AGN" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 3
        iLinie = 5
        iKunde = 4
        iAGN = 2
        iDatum = 6
        
    'AGN,Lieferant,Datum
    ElseIf cboSort3.Text = "Datum" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "AGN" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 3
        iLinie = 5
        iKunde = 6
        iAGN = 2
        iDatum = 4
        
    'AGN,Kunden,Lieferant
    ElseIf cboSort3.Text = "Lieferanten" And cboSort2.Text = "Kunden" And cboSort1.Text = "AGN" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 4
        iLinie = 5
        iKunde = 3
        iAGN = 2
        iDatum = 6
        
    'AGN,Kunden,Datum
    ElseIf cboSort3.Text = "Datum" And cboSort2.Text = "Kunden" And cboSort1.Text = "AGN" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Lieferanten"
        framegroesse4
        
        iLieferant = 5
        iLinie = 6
        iKunde = 3
        iAGN = 2
        iDatum = 4
        
    'AGN,Datum,Lieferant
    ElseIf cboSort3.Text = "Lieferanten" And cboSort2.Text = "Datum" And cboSort1.Text = "AGN" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Kunden"
        framegroesse4
        
        iLieferant = 4
        iLinie = 5
        iKunde = 6
        iAGN = 2
        iDatum = 3
        
    'AGN,Datum,Kunden
    ElseIf cboSort3.Text = "Kunden" And cboSort2.Text = "Datum" And cboSort1.Text = "AGN" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Lieferanten"
        framegroesse4
        
        iLieferant = 5
        iLinie = 6
        iKunde = 4
        iAGN = 2
        iDatum = 3
        
    'Kunden,AGN,Lieferanten
    ElseIf cboSort3.Text = "Lieferanten" And cboSort2.Text = "AGN" And cboSort1.Text = "Kunden" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 4
        iLinie = 5
        iKunde = 2
        iAGN = 3
        iDatum = 6
        
    'Kunden,AGN,Datum
    ElseIf cboSort3.Text = "Datum" And cboSort2.Text = "AGN" And cboSort1.Text = "Kunden" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Lieferanten"
        framegroesse4
        
        iLieferant = 5
        iLinie = 6
        iKunde = 2
        iAGN = 3
        iDatum = 4
        
    'Kunden,Lieferant,Linie
    ElseIf cboSort3.Text = "Linie" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "Kunden" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 3
        iLinie = 4
        iKunde = 2
        iAGN = 5
        iDatum = 6
        
    'Kunden,Lieferant,Datum
    ElseIf cboSort3.Text = "Datum" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "Kunden" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 3
        iLinie = 6
        iKunde = 2
        iAGN = 5
        iDatum = 4
        
    'Kunden,Lieferant,AGN
    ElseIf cboSort3.Text = "AGN" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "Kunden" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Datum"
        framegroesse4
        
        iLieferant = 3
        iLinie = 5
        iKunde = 2
        iAGN = 4
        iDatum = 6
        
    'Kunden,Datum,Lieferant
    ElseIf cboSort3.Text = "Lieferanten" And cboSort2.Text = "Datum" And cboSort1.Text = "Kunden" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "AGN"
        framegroesse4
        
        iLieferant = 4
        iLinie = 5
        iKunde = 2
        iAGN = 6
        iDatum = 3
        
    'Kunden,Datum,AGN
    ElseIf cboSort3.Text = "AGN" And cboSort2.Text = "Datum" And cboSort1.Text = "Kunden" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Lieferanten"
        framegroesse4
        
        iLieferant = 4
        iLinie = 5
        iKunde = 2
        iAGN = 6
        iDatum = 3
        
    'Datum,Lieferant,Linie
    ElseIf cboSort3.Text = "Linie" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "Datum" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Kunden"
        framegroesse4
        
        iLieferant = 3
        iLinie = 4
        iKunde = 6
        iAGN = 5
        iDatum = 2
        
    'Datum,Lieferant,Kunden
    ElseIf cboSort3.Text = "Kunden" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "Datum" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 3
        iLinie = 5
        iKunde = 4
        iAGN = 6
        iDatum = 2
        
    'Datum,Lieferant,AGN
    ElseIf cboSort3.Text = "AGN" And cboSort2.Text = "Lieferanten" And cboSort1.Text = "Datum" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Linie"
        cboSort4.AddItem "Kunden"
        framegroesse4
        
        iLieferant = 3
        iLinie = 5
        iKunde = 6
        iAGN = 4
        iDatum = 2
    
    'Datum,Kunden,Lieferant
    ElseIf cboSort3.Text = "Lieferanten" And cboSort2.Text = "Kunden" And cboSort1.Text = "Datum" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "AGN"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 4
        iLinie = 5
        iKunde = 3
        iAGN = 6
        iDatum = 2
        
    'Datum,Kunden,AGN
    ElseIf cboSort3.Text = "AGN" And cboSort2.Text = "Kunden" And cboSort1.Text = "Datum" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Lieferanten"
        framegroesse4
        
        iLieferant = 5
        iLinie = 6
        iKunde = 3
        iAGN = 4
        iDatum = 2
    
    'Datum,AGN,Lieferant
    ElseIf cboSort3.Text = "Lieferanten" And cboSort2.Text = "AGN" And cboSort1.Text = "Datum" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Kunden"
        cboSort4.AddItem "Linie"
        framegroesse4
        
        iLieferant = 4
        iLinie = 5
        iKunde = 6
        iAGN = 3
        iDatum = 2
        
    'Datum,AGN,Kunden
    ElseIf cboSort3.Text = "Kunden" And cboSort2.Text = "AGN" And cboSort1.Text = "Datum" Then
        cboSort4.Visible = True
        cboSort4.Clear
        cboSort4.Refresh
        cboSort4.AddItem "Lieferanten"
        framegroesse4
        
        iLieferant = 5
        iLinie = 6
        iKunde = 4
        iAGN = 3
        iDatum = 2
        
    ElseIf cboSort3.Text = "" Then
        cboSort2.Visible = False
        framegroesse3
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboSort3_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboSort4_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    caldate.Visible = False
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboSort4_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboSort3_GotFocus()
    On Error GoTo LOKAL_ERROR

    caldate.Visible = False
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboSort3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cboSort2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    caldate.Visible = False
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboSort2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


Private Sub cmdListDel_Click()
    On Error GoTo LOKAL_ERROR

    lstLinA.Clear
    
    If Frame3.Caption = "Linienauswahl" Then
        cboLin.Text = "alle Linien"
    ElseIf Frame3.Caption = "AGN-Auswahl" Then
        cboAgn.Text = "alle AGN´s"
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdListDel_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo LOKAL_ERROR
    
    
    If bQuick Then
        reportbildschirm "bedno", "aWKLata"
        
    ElseIf bQuickPlus Then
        reportbildschirm "bedpl", "aWKLatb"

    ElseIf bZ Then
        
        If optL.Value Then
            reportbildschirm "bedz", "aWKLatc"

        ElseIf optA.Value Then
            reportbildschirm "bedza", "aWKLatd"
        ElseIf optK.Value Then
            reportbildschirm "bedzk", "aWKLate"

        End If
        
    ElseIf bDetail Then
        reportbildschirm "bedple", "aWKLatf"
    Else
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPrint_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdQuick_Click()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    lblAnzeige.Caption = ""
    lblAnzeige.Refresh
    
    caldate.Visible = False
    MSHFLEX1.Visible = False
    
    If optq.Value = True Then
    
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        lstLinA.Clear
        
        ErstelleMSHFLEXq
        ErstelleSQLStatementq
        bQuick = True
    
        bQuickPlus = False
        bZ = False
        bDetail = False
        bsortOrAusw = False
        
    ElseIf optqp.Value = True Then
    
        Frame1.Visible = False
        Frame2.Visible = False
        Frame3.Visible = False
        lstLinA.Clear
        
        ErstelleMSHFLEXqPlus
        ErstelleSQLStatementqPlus
        bQuickPlus = True
        
        bQuick = False
        bZ = False
        bDetail = False
        bsortOrAusw = False
    
    ElseIf optz.Value = True Then
    
        ErstelleMSHFLEXZ
        ErstelleSQLStatementZ
        
        bZ = True
        
        bQuick = False
        bQuickPlus = False
        bDetail = False
        bFlexQ = False
       
    ElseIf optD.Value = True Then
    
        If Frame1.Visible = False And Frame2.Visible = False And Frame3.Visible = False Then
            iDatum = 2
            iLieferant = 3
            iLinie = 4
            iKunde = 5
            iAGN = 6
        End If
        
        ErstelleMSHFLEX
        ErstelleSQLStatement
        
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdQuick_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub cmdEnd_Click()
    On Error GoTo LOKAL_ERROR
    
   
            Unload frmWKLat
       
   
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ErstelleSQLStatement()
    On Error GoTo LOKAL_ERROR

    Dim Gr1 As String
    Dim Gr2 As String
    Dim Gr3 As String
    Dim Gr4 As String
    Dim Gr5 As String
    
    Dim sSQL As String
    Dim rs As Recordset
    Dim rsBed As Recordset
    Dim RsBedn As Recordset
    
    Dim lAnzahl, lrow As Long
    Dim counter As Integer
    Dim cFeld As String
    
    Dim cVon As String
    Dim cBis As String
    Dim lVon As Long
    Dim lBis As Long
    
    Dim sSQLlief As String
    Dim RsBez As Recordset
    Dim dLiNr As Double
    Dim sLief As String
   
    Dim sSQLAGN As String
    Dim rsagn As Recordset
    Dim dAgn As Double
    
    
    Dim sSQLKunde As String
    Dim sLin As String
   
    Dim sSQLLPZ As String
    Dim skun As String
    
    Dim dEkpr As Single
    Dim dEkWert As Single
    Dim iAnz As Integer
    Dim dPreis As Single
    Dim dUmsatz As Single
    Dim sTempBed As String
    Dim iBednu As Integer
    Dim sBedname As String
    Dim sbed As String
    Dim imon As Integer
    Dim iyear As Integer
    
    Dim iListanzahl As Integer
    Dim sT As String
    Dim i As Integer
    Dim sAGN As String
    Dim sLPZ As String
    
    Select Case iLieferant
        Case Is = 2
            Gr1 = "Linr"
        Case Is = 3
            Gr2 = "Linr"
        Case Is = 4
            Gr3 = "Linr"
        Case Is = 5
            Gr4 = "Linr"
        Case Is = 6
            Gr5 = "Linr"
        Case Else
            'nix tun
    End Select
    
    Select Case iLinie
        Case Is = 2
            Gr1 = "LPZ"
        Case Is = 3
            Gr2 = "LPZ"
        Case Is = 4
            Gr3 = "LPZ"
        Case Is = 5
            Gr4 = "LPZ"
        Case Is = 6
            Gr5 = "LPZ"
        Case Else
            'nix tun
    End Select
    
    Select Case iDatum
        Case Is = 2
            Gr1 = "adate"
        Case Is = 3
            Gr2 = "adate"
        Case Is = 4
            Gr3 = "adate"
        Case Is = 5
            Gr4 = "adate"
        Case Is = 6
            Gr5 = "adate"
        Case Else
            'nix tun
    End Select
    
    Select Case iAGN
        Case Is = 2
            Gr1 = "AGN"
        Case Is = 3
            Gr2 = "AGN"
        Case Is = 4
            Gr3 = "AGN"
        Case Is = 5
            Gr4 = "AGN"
        Case Is = 6
            Gr5 = "AGN"
        Case Else
            'nix tun
    End Select
    
    Select Case iKunde
        Case Is = 2
            Gr1 = "kundnr"
        Case Is = 3
            Gr2 = "kundnr"
        Case Is = 4
            Gr3 = "kundnr"
        Case Is = 5
            Gr4 = "kundnr"
        Case Is = 6
            Gr5 = "kundnr"
        Case Else
            'nix tun
    End Select
    
   
    lblAnzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblAnzeige.Refresh

    sBedname = Trim(cboBed.Text)
    
    If sBedname = "alle" Then
        sTempBed = ""
    Else
        sSQL = "Select Bednu from bedname"
        sSQL = sSQL & "  where bedname = '" & sBedname & "'"
        Set rsBed = gdBase.OpenRecordset(sSQL)
        
        If Not rsBed.EOF Then
        rsBed.MoveFirst
        
            If Not IsNull(rsBed!BEDNU) Then
                iBednu = rsBed!BEDNU
                sTempBed = " and BEDIENER = " & iBednu & ""
            Else
                Screen.MousePointer = 0
                lblAnzeige.Caption = "Bediener nicht gefunden"
                lblAnzeige.Refresh
                cboBed.SetFocus
                rsBed.Close: Set rsBed = Nothing
                Exit Sub
            End If
            
        Else
            Screen.MousePointer = 0
            lblAnzeige.Caption = "Bediener nicht gefunden"
            lblAnzeige.Refresh
            
            cboBed.SetFocus
            rsBed.Close: Set rsBed = Nothing
            Exit Sub
        End If
        rsBed.Close: Set rsBed = Nothing
    End If
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    Dim bLief As Boolean
    bLief = False
    
    If cboLief.Visible = True Then
        sLief = cboLief.Text
        If sLief <> "" Then
            bLief = True
            sSQL = "select Linr from lisrt where Liefbez = '" & sLief & "'"
            Set RsBez = gdBase.OpenRecordset(sSQL)
            If Not RsBez.EOF Then
                RsBez.MoveFirst
                If Not IsNull(RsBez!linr) Then
                    dLiNr = RsBez!linr
                    sSQLlief = " and LINR = " & dLiNr & " "
                Else
                    sSQLlief = ""
                End If
            End If
            RsBez.Close: Set RsBez = Nothing
        Else
            sSQLlief = ""
        End If
    End If
    
    Dim bkun As Boolean
    bkun = False
    
    If cboKunde.Visible = True Then
        skun = Trim(Right(cboKunde.Text, 6))
        
        
        If skun <> "Kunden" And skun <> "0" Then
         
            sSQLKunde = " and Kundnr = " & skun & " "
            skun = cboKunde.Text
            bkun = True
        Else
            sSQLKunde = ""
        End If
    Else
        sSQLKunde = ""
    End If
    
    loeschNEW "DETAIL", gdBase
    CreateTable "DETAIL", gdBase
    
    
    
    sSQL = "Insert into DETAIL Select Kassjour.Bediener, Kassjour.Linr, Kassjour.LPZ, Kassjour.AGN, Kassjour.KUNDNR, Kassjour.ARTNR, Kassjour.BEZEICH, Kassjour.Menge, Kassjour.Preis, Kassjour.EKPR, Kassjour.BELEGNR, Kassjour.MWST, Kassjour.adate, kassjour.BEZEICH as bedname, kassjour.preis as Ertrag, kassjour.preis as Ekwert, adate as mindat, adate as maxdat, kassjour.BEZEICH as auswahl "
'    sSQL = sSQL & " into detail "
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & sSQLlief
    sSQL = sSQL & sSQLKunde
    sSQL = sSQL & sTempBed
    
    sSQL = sSQL & " order BY Bediener, "
    sSQL = sSQL & " " & Gr1 & ", "
    sSQL = sSQL & " " & Gr2 & ", "
    sSQL = sSQL & " " & Gr3 & ", "
    sSQL = sSQL & " " & Gr4 & ", "
    sSQL = sSQL & " " & Gr5 & "  "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update detail set mindat = " & cVon & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update detail set maxdat = " & cBis & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update detail set auswahl = '" & sBedname & "' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update detail set auswahlLINR = '" & sLief & "' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    If lstLinA.Visible = True And lstLinA.ListCount > 0 Then
        If Frame3.Caption = "Linienauswahl" Then
    
            loeschNEW "GRIDA", gdBase
            CreateTable "GRIDA", gdBase
            sLPZ = ""
            iListanzahl = lstLinA.ListCount
            For i = 0 To iListanzahl - 1
                sT = lstLinA.list(i)
                If i = 0 Then
                    sSQL = "Insert into Grida Select Bediener, Linr, LPZ, AGN, KUNDNR, ARTNR, BEZEICH, Menge, Preis, EKPR, BELEGNR, MWST, adate, bedname, ertrag, ekwert, mindat, maxdat, auswahl "
                    sSQL = sSQL & "  From detail "
                    sSQL = sSQL & " where LPZ =  " & sT & ""
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sLPZ = sT & " " & ermLINBEZ(sT, CLng(dLiNr), gdBase)
                ElseIf i <> 0 Then
                    sSQL = "Insert into grida select * From detail "
                    sSQL = sSQL & " where LPZ =  " & sT & ""
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sLPZ = sLPZ & ", " & sT & " " & ermLINBEZ(sT, CLng(dLiNr), gdBase)
                End If
            Next i
        Else
            loeschNEW "GRIDA", gdBase
            CreateTable "GRIDA", gdBase


            sAGN = ""
            iListanzahl = lstLinA.ListCount
            For i = 0 To iListanzahl - 1
                sT = lstLinA.list(i)
                If sT <> "" Then
                    If i = 0 Then
                        sSQL = "Insert into grida Select Bediener, Linr, LPZ, AGN, KUNDNR, ARTNR, BEZEICH, Menge, Preis, EKPR, BELEGNR, MWST, adate, bedname, ertrag, ekwert, mindat, maxdat, auswahl"
                        sSQL = sSQL & "  From detail "
                        sSQL = sSQL & " where AGN =  " & sT & ""
                        gdBase.Execute sSQL, dbFailOnError
                        
                        sAGN = sT & " " & ermAGNbez(sT, gdBase)
                        
                    ElseIf i <> 0 Then
                        sSQL = "Insert into grida select * From detail "
                        sSQL = sSQL & " where AGN =  " & sT & ""
                        gdBase.Execute sSQL, dbFailOnError
                        
                        sAGN = sAGN & ", " & sT & " " & ermAGNbez(sT, gdBase)
                        
                    End If
                End If
            Next i
            
            
            
        End If
        
        loeschNEW "DETAIL", gdBase
        CreateTable "DETAIL", gdBase
        
        sSQL = "Insert into DETAIL select Bediener, Linr, LPZ, AGN, KUNDNR, ARTNR, BEZEICH, Menge, Preis, EKPR, BELEGNR, MWST, adate, BEZEICH as bedname, preis as Ertrag, preis as Ekwert,  mindat,  maxdat, auswahl,auswahlagn,auswahllinr,auswahllpz "
        sSQL = sSQL & " From grida "
        sSQL = sSQL & " order BY Bediener, "
        sSQL = sSQL & " " & Gr1 & ", "
        sSQL = sSQL & " " & Gr2 & ", "
        sSQL = sSQL & " " & Gr3 & ", "
        sSQL = sSQL & " " & Gr4 & ", "
        sSQL = sSQL & " " & Gr5 & "  "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    End If
    
    Set rs = gdBase.OpenRecordset("DETAIL", dbOpenTable)
    
    sSQL = "Update detail set mindat = " & cVon & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update detail set maxdat = " & cBis & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update detail set auswahl = '" & sBedname & "' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update detail set auswahlagn = '" & sAGN & "' "
    gdBase.Execute sSQL, dbFailOnError
            
    sSQL = "Update detail set auswahllpz = '" & sLPZ & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    counter = 0
    lrow = 0
    If Not rs.EOF Then
        
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!BEDIENER) Then
                sbed = rs!BEDIENER
            Else
                sbed = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sbed
            
            If cboBed.Text = "alle" Then

                sSQL = "Select Bedname from bedname"
                sSQL = sSQL & "  where bednu = " & sbed & ""
                Set RsBedn = gdBase.OpenRecordset(sSQL)

                If Not RsBedn.EOF Then
                    RsBedn.MoveFirst

                    If Not IsNull(RsBedn!BEDNAME) Then
                        sBedname = RsBedn!BEDNAME

                    Else
                        sBedname = "ohne Namen"
                    End If
                End If
                RsBedn.Close: Set RsBedn = Nothing

                MSHFLEX1.Col = 1
                MSHFLEX1.Text = sBedname
            Else
                MSHFLEX1.Col = 1
                MSHFLEX1.Text = sBedname
            End If
            
            rs.Edit
            rs!BEDNAME = sBedname
            rs.Update
            
            
            '********************variable spalten
            If Not IsNull(rs!Adate) Then
                cFeld = rs!Adate
            Else
                cFeld = ""
            End If

            MSHFLEX1.Col = iDatum
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!linr) Then
                cFeld = rs!linr
            Else
                cFeld = ""
            End If

            MSHFLEX1.Col = iLieferant
            MSHFLEX1.Text = cFeld



            If Not IsNull(rs!LPZ) Then
                cFeld = rs!LPZ
            Else
                cFeld = ""
            End If

            MSHFLEX1.Col = iLinie
            MSHFLEX1.Text = cFeld

            If Not IsNull(rs!AGN) Then
                cFeld = rs!AGN
            Else
                cFeld = ""
            End If

            MSHFLEX1.Col = iAGN
            MSHFLEX1.Text = cFeld

            If Not IsNull(rs!Kundnr) Then
                cFeld = rs!Kundnr
            Else
                cFeld = ""
            End If

            MSHFLEX1.Col = iKunde
            MSHFLEX1.Text = cFeld
            
            '***********feste Spalten
            
            If Not IsNull(rs!artnr) Then
                cFeld = rs!artnr
            Else
                cFeld = ""
            End If
    
            MSHFLEX1.Col = 7
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!BEZEICH) Then
                cFeld = rs!BEZEICH
            Else
                cFeld = ""
            End If
    
            MSHFLEX1.Col = 8
            MSHFLEX1.Text = cFeld
            
            
            If Not IsNull(rs!menge) Then
                iAnz = rs!menge
            Else
                iAnz = "0"
            End If
    
            MSHFLEX1.Col = 9
            MSHFLEX1.Text = iAnz
            

            If Not IsNull(rs!Preis) Then
                dPreis = rs!Preis
            Else
                dPreis = "0"
            End If

            MSHFLEX1.Col = 10
            cFeld = Format$(dPreis, "######0.00")
            MSHFLEX1.Text = cFeld

            
            If Not IsNull(rs!ekpr) Then
                dEkpr = rs!ekpr
            Else
                dEkpr = "0"
            End If

            MSHFLEX1.Col = 11
            cFeld = Format$(dEkpr, "######0.00")
            MSHFLEX1.Text = cFeld
            
            MSHFLEX1.Col = 12
            dEkWert = dEkpr * iAnz
            cFeld = Format$(dEkWert, "######0.00")
            MSHFLEX1.Text = cFeld
            
            rs.Edit
            rs!EKWERT = cFeld
            rs.Update
            
            MSHFLEX1.Col = 13
            dUmsatz = dPreis - dEkWert
            cFeld = Format$(dUmsatz, "######0.00")
            MSHFLEX1.Text = cFeld
            
            rs.Edit
            rs!ERTRAG = cFeld
            rs.Update
            
            If Not IsNull(rs!belegnr) Then
                cFeld = rs!belegnr
            Else
                cFeld = ""
            End If

            MSHFLEX1.Col = 14
            MSHFLEX1.Text = cFeld

            If Not IsNull(rs!MWST) Then
                cFeld = rs!MWST
            Else
                cFeld = ""
            End If

            MSHFLEX1.Col = 15
            MSHFLEX1.Text = cFeld
            
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    pbrZeit.Visible = False
    
    If bkun Then
        lblAnzeige.Caption = skun
        lblAnzeige.Refresh
    ElseIf bLief Then
        lblAnzeige.Caption = sLief
        lblAnzeige.Refresh
    Else
        lblAnzeige.Caption = "Detaildaten sind erstellt"
        lblAnzeige.Refresh
        
    End If
    
    If lrow = 0 Then
        lblAnzeige.Caption = "Keine Daten gefunden"
        lblAnzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
        optD.Value = False
    End If
    
    bDetail = True
    bQuick = False
    bQuickPlus = False
    bZ = False
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleSQLStatement"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    End If
    
End Sub
Private Sub ErstelleSQLStatementqPlus()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim rsKass As Recordset
    Dim rsrs As Recordset
    Dim rsBed As Recordset
    Dim RsBedn As Recordset
    Dim lAnzahl, lrow As Long
    Dim lAktRecord As Long
    Dim counter As Integer
    Dim cFeld As String
    
    Dim cVon As String
    Dim cBis As String
    Dim lVon As Long
    Dim lBis As Long
    
    Dim iAnzahl As Integer
    Dim dEkpr As Single
    Dim dEkWert As Single
    Dim iAnz As Integer
    Dim dPreis As Single
    Dim dErtrag As Single
    Dim sTempBed As String
    Dim iBednu As Integer
    Dim sBedname As String
    Dim sbed As String
    Dim dwert As Integer
    Dim iAnzahlKunden As Integer
    Dim dProzent As Single
    Dim dGesamtumsatz As Single
   
    lblAnzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblAnzeige.Refresh

    sBedname = Trim(cboBed.Text)
    
    If sBedname = "alle" Then
        sTempBed = ""
    Else
        sSQL = "Select Bednu from bedname"
        sSQL = sSQL & " where bedname = '" & sBedname & "'"
        Set rsBed = gdBase.OpenRecordset(sSQL)
        
        If Not rsBed.EOF Then
        rsBed.MoveFirst
        
            If Not IsNull(rsBed!BEDNU) Then
                iBednu = rsBed!BEDNU
                sTempBed = " and Kassjour.BEDIENER = " & iBednu & ""
            Else
                sTempBed = ""
            End If
        End If
        rsBed.Close: Set rsBed = Nothing
           
        
        
    End If
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))

    sSQL = "Drop Table bedpl"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select KASSJOUR.Bediener, BEDNAME.BEDNAME , Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis , min(adate) as mindat, max(adate) as maxdat, min(azeit) as auswahl, count(kassjour.BELEGNR) as ANZKUNDEN , min (Preis)as ERTRAG, min (Preis)as KUSCHNITT "
    sSQL = sSQL & " into Bedpl"
    sSQL = sSQL & " from Kassjour, BEDNAME"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & sTempBed
    
    sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedpl set mindat = " & cVon & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedpl set maxdat = " & cBis & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedpl set auswahl = '" & sBedname & "' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Set rs = gdBase.OpenRecordset("Bedpl", dbOpenTable)
    
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    
    counter = 0
    lrow = 0
    If Not rs.EOF Then
        
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!BEDIENER) Then
                sbed = rs!BEDIENER
            Else
                sbed = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sbed
            
            If Not IsNull(rs!BEDNAME) Then
            cFeld = rs!BEDNAME
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!ANZAHL) Then
                iAnz = rs!ANZAHL
            Else
                iAnz = "0"
            End If
    
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = iAnz
            

            If Not IsNull(rs!UMSATZ) Then
                dPreis = rs!UMSATZ
            Else
                dPreis = "0"
            End If
            
            dGesamtumsatz = dGesamtumsatz + dPreis
            
            MSHFLEX1.Col = 3
            cFeld = Format$(dPreis, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            If Not IsNull(rs!EinKPreis) Then
                dEkpr = rs!EinKPreis
            Else
                dEkpr = "0"
            End If

            MSHFLEX1.Col = 4
            cFeld = Format$(dEkpr, "######0.00")
            MSHFLEX1.Text = cFeld
            

            MSHFLEX1.Col = 5
            dErtrag = dPreis - dEkpr
            cFeld = Format$(dErtrag, "######0.00")
            MSHFLEX1.Text = cFeld
            
            rs.Edit
            rs!ERTRAG = cFeld
            rs.Update
    
            'Anzahl Kunden
            'zuerst kassenzahl abfragen
            
            sSQL = "Select distinct Kasnum "
            sSQL = sSQL & "from Kassjour "
            sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
            Set rsKass = gdBase.OpenRecordset(sSQL)
            If Not rsKass.EOF Then
            
                rsKass.MoveFirst
                Do While Not rsKass.EOF
                
                    If Not IsNull(rsKass!KASNUM) Then
                    
                        
                        sSQL = "Select  distinct adate, BELEGNR as ANZKUNDEN "
                        sSQL = sSQL & " from Kassjour "
                        sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
                        sSQL = sSQL & " and bediener = " & sbed & " "
                        sSQL = sSQL & " and Kassjour.artnr <> 666666 "
                        sSQL = sSQL & " and Kassjour.kasnum = " & rsKass!KASNUM
                        
                        Set rsrs = gdBase.OpenRecordset(sSQL)
                        
                        If Not rsrs.EOF Then
                            iAnzahlKunden = iAnzahlKunden + rsrs.RecordCount
                        Else
                            iAnzahlKunden = iAnzahlKunden + 0
                        End If
                        rsrs.Close: Set rsrs = Nothing
                    End If
                rsKass.MoveNext
                Loop
            
            End If
            rsKass.Close
            
            
            
            rs.Edit
            rs!ANZKUNDEN = iAnzahlKunden
            
            rs.Update
    
            cFeld = Format$(iAnzahlKunden, "###,###,##0")
            MSHFLEX1.Col = 8
            MSHFLEX1.Text = cFeld
            
            
            If dPreis <> 0 And iAnzahlKunden <> 0 Then
                dPreis = dPreis / iAnzahlKunden
            Else
                dPreis = 0
            End If
            
            rs.Edit
            rs!Kuschnitt = dPreis
            rs.Update
    
            
            MSHFLEX1.Col = 7
            MSHFLEX1.Text = Format$(dPreis, "##0.00")
            
                
            If iAnzahlKunden <> 0 Then
                dPreis = iAnz / iAnzahlKunden
            Else
                dPreis = 0
            End If
            
            MSHFLEX1.Col = 9
            MSHFLEX1.Text = Format$(dPreis, "##0.00")
            
            iAnzahlKunden = 0
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    
    For lAktRecord = 1 To lrow
        MSHFLEX1.Row = lAktRecord
        MSHFLEX1.Col = 3
        cFeld = MSHFLEX1.Text

        cFeld = fnMoveComma2Point$(cFeld)
        dProzent = Val(cFeld)

        dProzent = dProzent * 100
        If dGesamtumsatz <> 0 Then
            dProzent = dProzent / dGesamtumsatz
        Else
            dProzent = 100
        End If
        MSHFLEX1.Col = 6
        MSHFLEX1.Text = Format$(dProzent, "##0.00")


    Next lAktRecord
    
    pbrZeit.Visible = False
    lblAnzeige.Caption = "erweiterte Übersicht erstellt"
    lblAnzeige.Refresh
    
    If lrow = 0 Then
        lblAnzeige.Caption = "Keine Daten gefunden"
        lblAnzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ErstelleSQLStatementqPlus"
        Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Integer
    Dim i           As Integer
    Dim j           As Integer
    
    With gridx
    
        ReDim bBreit(.Cols - 1)
        
        For j = 0 To .Rows - 1
            For i = 0 To .Cols - 1
                If TextWidth(.TextMatrix(j, i)) > bBreit(i) Then
                    bBreit(i) = TextWidth(.TextMatrix(j, i))
                End If
            Next i
        Next j
        
        
        Select Case Screen.Height
            Case Is > 15000
                siFak = 1.5
            Case Is > 12000
                siFak = 1.4
            Case Is > 11000
                siFak = 1.2
            Case Is > 10000
                siFak = 1.1
            Case Is > 8000
                siFak = 1#
        End Select
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = bBreit(i) * siFak * siEigFak
        Next i
    
    End With
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErstelleSQLStatementq()
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rs          As Recordset
    Dim rsBed       As Recordset
    Dim lAnzahl     As Long
    Dim lrow        As Long
    Dim counter     As Integer
    Dim cFeld       As String
    
    Dim cVon        As String
    Dim cBis        As String
    Dim lVon        As Long
    Dim lBis        As Long
    
    Dim iAnzahl     As Integer
    Dim dEkpr       As Single
    Dim dEkWert     As Single
    Dim iAnz        As Long
    Dim dPreis      As Single
    Dim dErtrag     As Single
    Dim sTempBed    As String
    Dim iBednu      As Integer
    Dim sBedname    As String
    Dim sbed        As String
    Dim dwert       As Integer
   
   
    lblAnzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblAnzeige.Refresh

    sBedname = Trim(cboBed.Text)
    
    If sBedname = "alle" Then
        sTempBed = ""
    Else
        sSQL = "Select Bednu from bedname"
        sSQL = sSQL & "  where bedname = '" & sBedname & "'"
        Set rsBed = gdBase.OpenRecordset(sSQL)
        
        If Not rsBed.EOF Then
            rsBed.MoveFirst
        
            If Not IsNull(rsBed!BEDNU) Then
                iBednu = rsBed!BEDNU
                sTempBed = " and Kassjour.BEDIENER = " & iBednu & ""
            Else
                sTempBed = ""
            End If
        End If
        rsBed.Close: Set rsBed = Nothing
    End If
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "bedno", gdBase
    CreateTable "BEDNO", gdBase
    
    sSQL = "Insert into Bedno Select KASSJOUR.Bediener, BEDNAME.BEDNAME , Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis, min(adate) as mindat, max(adate) as maxdat, min(azeit) as auswahl"
'    sSQL = sSQL & " into Bedno"
    sSQL = sSQL & " from Kassjour, BEDNAME"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & sTempBed
    sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedno set mindat = " & cVon & ""
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedno set maxdat = " & cBis & ""
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedno set auswahl = '" & sBedname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Update Bedno set auswahlLinr = '" & sBedname & "' "
'    gdbase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update Bedno set auswahlLpz = '" & sBedname & "' "
'    gdbase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update Bedno set auswahlAGN = '" & sBedname & "' "
'    gdbase.Execute sSQL, dbFailOnError
    
    Set rs = gdBase.OpenRecordset("Bedno", dbOpenTable)
    
    If Not rs.EOF Then
        rs.MoveLast
        
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    counter = 0
    lrow = 0
    If Not rs.EOF Then
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!BEDIENER) Then
                sbed = rs!BEDIENER
            Else
                sbed = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sbed
            
            If Not IsNull(rs!BEDNAME) Then
            cFeld = rs!BEDNAME
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            If Not IsNull(rs!ANZAHL) Then
                iAnz = rs!ANZAHL
            Else
                iAnz = "0"
            End If
    
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = iAnz
            

            If Not IsNull(rs!UMSATZ) Then
                dPreis = rs!UMSATZ
            Else
                dPreis = "0"
            End If
            
            MSHFLEX1.Col = 3
            cFeld = Format$(dPreis, "######0.00")
            MSHFLEX1.Text = cFeld
            
            
            If Not IsNull(rs!EinKPreis) Then
                dEkpr = rs!EinKPreis
            Else
                dEkpr = "0"
            End If

            MSHFLEX1.Col = 4
            cFeld = Format$(dEkpr, "######0.00")
            MSHFLEX1.Text = cFeld
            

            MSHFLEX1.Col = 5
            dErtrag = dPreis - dEkpr
            cFeld = Format$(dErtrag, "######0.00")
            MSHFLEX1.Text = cFeld
            
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
            
           
    pbrZeit.Visible = False
    lblAnzeige.Caption = "einfache Übersicht erstellt"
    lblAnzeige.Refresh
    
    If lrow = 0 Then
        lblAnzeige.Caption = "Keine Daten gefunden"
        lblAnzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
        bFlexQ = True
    End If
  
Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ErstelleSQLStatementq"
        Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub
Private Sub ErstelleMSHFLEX()
    On Error GoTo LOKAL_ERROR
    
'    MSHFLEX1.Height = 5895
'    MSHFLEX1.Left = 480
'    MSHFLEX1.Top = 960
'    MSHFLEX1.Width = 10815
'
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 16
        .FixedCols = 1
        
        '***feste Spalten****
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 600
        .Text = "Bed.Nr"
        
        .Col = 1
        .ColWidth(1) = 2000
        .Text = "Bedienername"
        
        '****************************************
        '*  variable Spalten                    *
        '*  Verschiebung durch Group - Klausel  *
        '****************************************
        
        .Col = iLieferant
        .ColWidth(iLieferant) = 800
        .Text = "Lieferant"
        
        .Col = iLinie
        .ColWidth(iLinie) = 800
        .Text = "Linie"
        
        .Col = iKunde
        .ColWidth(iKunde) = 800
        .Text = "Kunden"
        
        .Col = iAGN
        .ColWidth(iAGN) = 800
        .Text = "AGN"
        
        .Col = iDatum
        .ColWidth(iDatum) = 800
        .Text = "Datum"
        '****************************************
        '*  feste Spalten                       *
        '****************************************
        
        .Col = 7
        .ColWidth(7) = 700
        .Text = "ArtNr"
        
        .Col = 8
        .ColWidth(8) = 3500
        .Text = "Artikelbezeichnung"
        
        .Col = 9
        .ColWidth(9) = 600
        .Text = "Anzahl"
        
        .Col = 10
        .ColWidth(10) = 700
        .Text = "Umsatz"
        
        .Col = 11
        .ColWidth(11) = 700
        .Text = "EK Preis"
        
        .Col = 12
        .ColWidth(12) = 700
        .Text = "EK Wert"
        
        .Col = 13
        .ColWidth(13) = 700
        .Text = "Ertrag"
        
        .Col = 14
        .ColWidth(14) = 700
        .Text = "BelegNr"
        
        .Col = 15
        .ColWidth(15) = 600
        .Text = "MWST"
    
    End With
    

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEX"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ErstelleMSHFLEXqPlus()
    On Error GoTo LOKAL_ERROR
    
'    MSHFLEX1.Height = 5895
'    MSHFLEX1.Left = 480
'    MSHFLEX1.Top = 960
'    MSHFLEX1.Width = 10815
    
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 10
        .FixedCols = 2
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 600
        .Text = "Bed.Nr"
        
        .Col = 1
        .ColWidth(1) = 2000
        .Text = "Bedienername"
        
        .Col = 2
        .ColWidth(2) = 1000
        .Text = "verk.Artikel"
        
        .Col = 3
        .ColWidth(3) = 1000
        .Text = "Umsatz(VK)"
        
        .Col = 4
        .ColWidth(4) = 1000
        .Text = "Umsatz(EK)"
        
        .Col = 5
        .ColWidth(5) = 800
        .Text = "Ertrag"
        
        .Col = 6
        .ColWidth(6) = 1200
        .Text = "%Anteil Umsatz"
        
        .Col = 7
        .ColWidth(7) = 1300
        .Text = "Kundenschnitt €"
        
        .Col = 8
        .ColWidth(8) = 1200
        .Text = "Anzahl Kunden"
        
        .Col = 9
        .ColWidth(9) = 1800
        .Text = "verk. Artikel pro Kunde"
        
    End With
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXqPlus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ErstelleMSHFLEXq()
    On Error GoTo LOKAL_ERROR
    
'    MSHFLEX1.Height = 5895
'    MSHFLEX1.Left = 480
'    MSHFLEX1.Top = 960
'    MSHFLEX1.Width = 6775
'
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 6
        .FixedCols = 1
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 600
        .Text = "Bed.Nr"
        
        .Col = 1
        .ColWidth(1) = 2000
        .Text = "Bedienername"
        
        .Col = 2
        .ColWidth(2) = 1000
        .Text = "verk.Artikel"
        
        .Col = 3
        .ColWidth(3) = 1000
        .Text = "Umsatz(VK)"
        
        .Col = 4
        .ColWidth(4) = 1000
        .Text = "Umsatz(EK)"
        
        .Col = 5
        .ColWidth(5) = 800
        .Text = "Ertrag"
        
       
    
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXq"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErstelleMSHFLEXZ()
    On Error GoTo LOKAL_ERROR
    
    Dim sVar As String
    
'    MSHFLEX1.Height = 5895
'    MSHFLEX1.Left = 480
'    MSHFLEX1.Top = 960
'    MSHFLEX1.Width = 8775
    
    
        If optK.Value = True Then
            sVar = "Kunden"
        ElseIf optA.Value = True Then
            sVar = "AGN"
        ElseIf optL.Value = True Then
            sVar = "Lieferant"
        End If

    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 7
        .FixedCols = 2
   
        .Row = 0
        .Col = 0
        .ColWidth(0) = 600
        .Text = "Bed.Nr"
        
        .Col = 1
        .ColWidth(1) = 2000
        .Text = "Bedienername"
        
        .Col = 2
        .ColWidth(2) = 2000
        .Text = sVar
        
        .Col = 3
        .ColWidth(3) = 1000
        .Text = "verk.Artikel"
        
        .Col = 4
        .ColWidth(4) = 1000
        .Text = "Umsatz(VK)"
        
        .Col = 5
        .ColWidth(5) = 1000
        .Text = "Umsatz(EK)"
        
        .Col = 6
        .ColWidth(6) = 800
        .Text = "Ertrag"
        
       
    
    End With
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleMSHFLEXZ"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ErstelleSQLStatementZ()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim rsBed As Recordset
    Dim lAnzahl, lrow As Long
    Dim counter As Integer
    Dim cFeld As String
    
    Dim cVon As String
    Dim cBis As String
    Dim lVon As Long
    Dim lBis As Long
    
    Dim iAnzahl As Integer
    Dim lLief As Long
    Dim dEkpr As Single
    Dim dEkWert As Single
    Dim iAnz As Integer
    Dim dPreis As Single
    Dim dErtrag As Single
    Dim sTempBed As String
    Dim iBednu As Integer
    Dim sBedname As String
    Dim sbed As String
    Dim dwert As Integer
    Dim sSelect As String
    Dim sGROUP As String
    Dim lVari As Long
   
    lblAnzeige.Caption = "Daten für diesen Zeitraum werden ermittelt..."
    lblAnzeige.Refresh

    sBedname = Trim(cboBed.Text)
    
    If sBedname = "alle" Then
        sTempBed = ""
    Else
        sSQL = "Select Bednu from bedname"
        sSQL = sSQL & "  where bedname = '" & sBedname & "'"
        Set rsBed = gdBase.OpenRecordset(sSQL)
        
        If Not rsBed.EOF Then
        rsBed.MoveFirst
        
            If Not IsNull(rsBed!BEDNU) Then
                iBednu = rsBed!BEDNU
                sTempBed = " and Kassjour.BEDIENER = " & iBednu & ""
            Else
                sTempBed = ""
            End If
        End If
        rsBed.Close: Set rsBed = Nothing
    End If
           
           
    Dim lLinr As Long
    Dim sLieferant As String
    Dim sSQLLieferant As String
    Dim rslief As Recordset
    
    Dim sSQLKunden As String
    Dim sKUNDEN As String
    
    Dim sAGN As String
    Dim sSQLAGN As String
    
    cVon = Text1(0).Text
    cBis = Text1(1).Text
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)
    
    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "bedz", gdBase
    
    
    
    If cboSort1.Visible = True Then
    
        If cboSort1.Text = "Lieferanten" Then
            sSelect = ",Kassjour.LINR"
            sGROUP = "LINR"
            
            sLieferant = ""
            sLieferant = Trim(cboLief.Text)
            
            If sLieferant = "alle Lieferanten" Then
                sSQLLieferant = ""
            Else
                sSQL = "Select LINR from lisrt"
                sSQL = sSQL & "  where Liefbez = '" & sLieferant & "'"
                Set rslief = gdBase.OpenRecordset(sSQL)

                If Not rslief.EOF Then
                    rslief.MoveFirst

                    If Not IsNull(rslief!linr) Then
                        lLinr = rslief!linr
                        sSQLLieferant = " and Kassjour.LINR = " & lLinr & ""
                    Else
                        sSQLLieferant = ""
                    End If
                End If
                rslief.Close
            End If
            
    
        ElseIf cboSort1.Text = "Kunden" Then
            sSelect = ",Kassjour.KUNDNR"
            sGROUP = "KUNDNR"
            
            sKUNDEN = ""
            sKUNDEN = Trim(Right(cboKunde.Text, 6))
            
            If sKUNDEN = "Kunden" Then
                sSQLKunden = ""
            Else
                sSQLKunden = " and Kassjour.KUNDNR = " & sKUNDEN & ""
            End If
            
            
            
            
        ElseIf cboSort1.Text = "AGN" Then
            sSelect = ",Kassjour.AGN"
            sGROUP = "AGN"
            
            sAGN = ""
            sAGN = Trim(Right(cboAgn.Text, 3))
            
            If sAGN = "N´s" Then
                sSQLAGN = ""
            Else
                sSQLAGN = " and Kassjour.AGN = " & sAGN & ""
            End If

        ElseIf cboSort1.Text = "Datum" Then
            sSelect = ",Kassjour.LINR"
            sGROUP = "LINR"
        Else
            sSelect = ""
        End If
        
    Else
        
        sSelect = ",Kassjour.LINR"
        sGROUP = "LINR"
    End If
    

    If optK.Value Then
        sSelect = ",Kassjour.KUNDNR"
        sGROUP = "KUNDNR"
        
        sSQL = "Select KASSJOUR.Bediener, BEDNAME.BEDNAME, Kunden.Name  " & sSelect & " ,  Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis, min(adate) as mindat, max(adate) as maxdat, min(azeit) as auswahl"
        sSQL = sSQL & " into Bedz"
        sSQL = sSQL & " from Kassjour, BEDNAME, Kunden"
        sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
        sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
        sSQL = sSQL & " and kunden.kundnr = kassjour.kundnr "
        sSQL = sSQL & " and Kassjour.artnr <> 666666 "
        sSQL = sSQL & sTempBed
        sSQL = sSQL & sSQLLieferant
        sSQL = sSQL & sSQLKunden
        sSQL = sSQL & sSQLAGN
        sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME " & sSelect & " ,Kunden.Name "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        
    ElseIf optA.Value Then
        sSelect = ",Kassjour.AGN"
        sGROUP = "AGN"
        
        sSQL = "Select KASSJOUR.Bediener, BEDNAME.BEDNAME, AGNDBF.AGTEXT  " & sSelect & " ,  Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis, min(adate) as mindat, max(adate) as maxdat, min(azeit) as auswahl"
        sSQL = sSQL & " into Bedz"
        sSQL = sSQL & " from Kassjour, BEDNAME, AGNDBF"
        sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
        sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
        sSQL = sSQL & " and AGNDBF.AGN = kassjour.AGN "
        sSQL = sSQL & " and Kassjour.artnr <> 666666 "
        sSQL = sSQL & sTempBed
        sSQL = sSQL & sSQLLieferant
        sSQL = sSQL & sSQLKunden
        sSQL = sSQL & sSQLAGN
        sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME " & sSelect & " , AGNDBF.AGTEXT "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        
        
    Else
    
    sSQL = "Select KASSJOUR.Bediener, BEDNAME.BEDNAME,lisrt.liefbez  " & sSelect & " ,  Sum(Menge)as Anzahl, Sum (Preis)as Umsatz, Sum(EKPR*Menge)as EinKPreis, min(adate) as mindat, max(adate) as maxdat, min(azeit) as auswahl"
    sSQL = sSQL & " into Bedz"
    sSQL = sSQL & " from Kassjour, BEDNAME, lisrt"
    sSQL = sSQL & " Where Kassjour.ADATE Between " & cVon & " And " & cBis & " "
    sSQL = sSQL & " and Kassjour.bediener = bedname.bednu "
    sSQL = sSQL & " and Kassjour.artnr <> 666666 "
    sSQL = sSQL & " and lisrt.linr = kassjour.linr "
    sSQL = sSQL & sTempBed
    sSQL = sSQL & sSQLLieferant
    sSQL = sSQL & sSQLKunden
    sSQL = sSQL & sSQLAGN
    sSQL = sSQL & " group BY  KASSJOUR.Bediener, BEDNAME.BEDNAME " & sSelect & " ,lisrt.liefbez "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    End If

    
    sSQL = "Update Bedz set mindat = " & cVon & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedz set maxdat = " & cBis & ""
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Bedz set auswahl = '" & sBedname & "' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Set rs = gdBase.OpenRecordset("Bedz", dbOpenTable)
    
    If Not rs.EOF Then
        rs.MoveLast
        lAnzahl = rs.RecordCount
        pbrZeit.Visible = True
        pbrZeit.Max = lAnzahl
        rs.MoveFirst
    End If
    
    counter = 0
    lrow = 0
    If Not rs.EOF Then
        
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            
            counter = counter + 1
            pbrZeit.Value = counter
            
            lrow = lrow + 1
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            If Not IsNull(rs!BEDIENER) Then
                sbed = rs!BEDIENER
            Else
                sbed = "00000"
            End If
    
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = sbed
            
            If Not IsNull(rs!BEDNAME) Then
            cFeld = rs!BEDNAME
            Else
                cFeld = ""
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = Trim(cFeld)
            
            Select Case sGROUP
                Case Is = "AGN"
            
                If Not IsNull(rs!AGN) Then
                    lVari = rs!AGN
                Else
                    lVari = "0"
                End If
        
                MSHFLEX1.Col = 2
                MSHFLEX1.Text = lVari
                
            Case Is = "LINR"
                
                If Not IsNull(rs!linr) Then
                    lVari = rs!linr
                Else
                    lVari = "0"
                End If
        
                MSHFLEX1.Col = 2
                MSHFLEX1.Text = lVari
                
             Case Is = "KUNDNR"
                
                If Not IsNull(rs!Kundnr) Then
                    lVari = rs!Kundnr
                Else
                    lVari = "0"
                End If
        
                MSHFLEX1.Col = 2
                MSHFLEX1.Text = lVari
                Case Else
            'nix tun
            End Select
            
            If Not IsNull(rs!ANZAHL) Then
                iAnz = rs!ANZAHL
            Else
                iAnz = "0"
            End If
    
            MSHFLEX1.Col = 3
            MSHFLEX1.Text = iAnz
            
            If Not IsNull(rs!UMSATZ) Then
                dPreis = rs!UMSATZ
            Else
                dPreis = "0"
            End If
            
            MSHFLEX1.Col = 4
            cFeld = Format$(dPreis, "######0.00")
            MSHFLEX1.Text = cFeld
            
            If Not IsNull(rs!EinKPreis) Then
                dEkpr = rs!EinKPreis
            Else
                dEkpr = "0"
            End If

            MSHFLEX1.Col = 5
            cFeld = Format$(dEkpr, "######0.00")
            MSHFLEX1.Text = cFeld
            
            MSHFLEX1.Col = 6
            dErtrag = dPreis - dEkpr
            cFeld = Format$(dErtrag, "######0.00")
            MSHFLEX1.Text = cFeld
            
             rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
            
    pbrZeit.Visible = False
    lblAnzeige.Caption = "Zusammenfassung erstellt"
    lblAnzeige.Refresh
    
    If lrow = 0 Then
        lblAnzeige.Caption = "Keine Daten gefunden"
        lblAnzeige.Refresh
        MSHFLEX1.Visible = False
    Else
        MSHFLEX1.Visible = True
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
        bFlexQ = True
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErstelleSQLStatementZ"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub WerteEingabenAus()
    On Error GoTo LOKAL_ERROR

    Dim sBediener As String
    
    Dim vonDat As String
    Dim bisDat As String
    
    Dim sSort1 As String
    Dim sSort2 As String
    Dim sSort3 As String
    Dim sSort4 As String
    
    Dim sLieferant As String
    Dim sLinie As String
    Dim sAGN As String
    Dim sKunde As String
    
    
    
    
    sBediener = cboBed.Text
    vonDat = Text1(0).Text
    bisDat = Text1(1).Text
    
    sSort1 = cboSort1.Text
    sSort2 = cboSort2.Text
    sSort3 = cboSort3.Text
    sSort4 = cboSort4.Text
    
    
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WerteEingabenAus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub cmdZusammenstellen_Click()
    On Error GoTo LOKAL_ERROR
    
    lblAnzeige.Caption = ""
    lblAnzeige.Refresh
    
    framegroesse1
    Frame3.Visible = False
    lstLinA.Clear
    Frame2.Visible = False
    caldate.Visible = False
    MSHFLEX1.Visible = False
    cboSort1.SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdZusammenstellen_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub cmdDel_Click()
    On Error GoTo LOKAL_ERROR
    
    optq.Value = True
    optL.Value = True
    
    lblAnzeige.Caption = ""
    lblAnzeige.Refresh
        
    Frame1.Visible = False
    cboBed.Text = "alle"
    cboBed.SetFocus
    
    
    Frame2.Visible = False
    caldate.Visible = False
    MSHFLEX1.Visible = False
    bLinie = False
    
    Frame3.Visible = False
    lstLinA.Clear
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdDel_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub füllecboKunden()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim sTemp As String
    Dim sSatz As String
    Dim counter As Long
    Dim lAnzahl As Long
    
    Set rs = gdBase.OpenRecordset("KUNDEN", dbOpenTable)
    
    If Not rs.EOF Then
        rs.MoveLast
        lAnzahl = rs.RecordCount

        
        rs.MoveFirst

        cboKunde.Clear
        Do While Not rs.EOF
            
            If counter = 2000 Then
                counter = 0
            End If
            counter = counter + 1

        
            If Not IsNull(rs!name) Then
                sTemp = rs!name
                sTemp = Trim(sTemp)
                sTemp = sTemp & Space(2)
                sSatz = sTemp
            Else
                sTemp = ""
            End If
            
            If Not IsNull(rs!vorname) Then
                sTemp = rs!vorname
            Else
                sTemp = ""
            End If
            
            sTemp = Trim(sTemp)
            sTemp = sTemp & Space(2)
            sSatz = sSatz & sTemp & Space(4)
            
            If Not IsNull(rs!Kundnr) Then
                sTemp = rs!Kundnr
            Else
                sTemp = ""
            End If
            
            sTemp = Trim(sTemp)
            sSatz = sSatz & sTemp
            
            cboKunde.AddItem UCase(sSatz)
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllecboKunden"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub fülledat()
    On Error GoTo LOKAL_ERROR

    cbodat.Clear
    cbodat.AddItem "Zeitraum auswählen"
    cbodat.AddItem "Heute"
    cbodat.AddItem "Gestern"
    cbodat.AddItem "letzten 3 Tage"
    cbodat.AddItem "letzte Woche"
    cbodat.AddItem "letzter Monat"
    cbodat.AddItem "dieser Monat"
    cbodat.AddItem "letzten 30 Tage"
    
    
    
    
    
   
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fülledat"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub füllecboLieferanten()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim sTemp As String
    Dim sSatz As String
    Dim counter As Long
    Dim lAnzahl As Long
    
    sSQL = "select liefbez from lisrt order by liefbez"
    Set rs = gdBase.OpenRecordset(sSQL)
    
    cboLief.Clear
    If Not rs.EOF Then
        rs.MoveFirst

        
        Do While Not rs.EOF
        
            If Not IsNull(rs!LIEFBEZ) Then
                sTemp = Trim(rs!LIEFBEZ)
            Else
                sTemp = ""
            End If
            
           
            If sTemp <> "" Then
                cboLief.AddItem sTemp
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close: Set rs = Nothing
    cboLief.Text = "alle Lieferanten"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllecboLieferanten"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_Click()
    On Error GoTo LOKAL_ERROR
    
    caldate.Visible = False
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
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
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
        
    Screen.MousePointer = 11
    
    positionierenwklat
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
        
    iDatum = 2
    iLieferant = 3
    iLinie = 4
    iKunde = 5
    iAGN = 6
    
    optq.Value = True
    
    bQuick = False
    bQuickPlus = False
    bZ = False
    bDetail = False
    bFlexQ = False
    bsortOrAusw = False
    
    gbNoData = False
    
   
    
    füllecboBediener cboBed
    

    
    caldate.Value = Date
    
    
    bLinie = False
    
    Text1(1).Text = Date
    Text1(0).Text = Date - 7
    
    Screen.MousePointer = Default
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub positionierenwklat()
    On Error GoTo LOKAL_ERROR
    
    caldate.Height = 3015
    caldate.Left = 3240
    caldate.Top = 2960
    caldate.Width = 5055
    
    MSHFLEX1.Height = 5895
    MSHFLEX1.Left = 480
    MSHFLEX1.Top = 960
    MSHFLEX1.Width = 10815
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionierenwklat"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub



Private Sub lstLinA_Click()
    On Error GoTo LOKAL_ERROR

     lstLinA.RemoveItem (lstLinA.ListIndex)
     

     If lstLinA.ListCount = 0 Then
        If Frame3.Caption = "AGN-Auswahl" Then
            cboAgn.Text = "alle AGN´s"
        ElseIf Frame3.Caption = "Linienauswahl" Then
            cboLin.Text = "alle Linien"
        End If
     End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstLinA_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub



Private Sub mnuBed_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        
        Case Is = 0 'KUCUT Bediener
            schreibeProtokollProgrammablauf " löst Liste aus    " & mnuBed(Index).Caption
            BestBedKuCut txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & mnuBed(Index).Caption
        Case Is = 1 'KUCUT Bediener Entwicklung
            schreibeProtokollProgrammablauf " löst Liste aus    " & mnuBed(Index).Caption
            BestBedKuCutDEVELo txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & mnuBed(Index).Caption
        Case Is = 2
            schreibeProtokollProgrammablauf " löst Liste aus    " & mnuBed(Index).Caption
            BestBedKuCut1 txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & mnuBed(Index).Caption
        Case Is = 3
            schreibeProtokollProgrammablauf " löst Liste aus    " & mnuBed(Index).Caption
            BestBedKuCut2 txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & mnuBed(Index).Caption
        Case Is = 4
            schreibeProtokollProgrammablauf " löst Liste aus    " & mnuBed(Index).Caption
            BestBedProvision txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & mnuBed(Index).Caption
        Case Is = 5
            schreibeProtokollProgrammablauf " löst Liste aus    " & mnuBed(Index).Caption
            BestBedProvisionRab txtStatus, picprogress
            schreibeProtokollProgrammablauf " Liste fertig      " & mnuBed(Index).Caption
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "mnuBed_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim sBdname As String
    
    sBdname = ""
    
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MSHFLEX1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim lcol As Long
    Dim lrow As Long
    Dim sBdname As String
    
    sBdname = ""
    lcol = MSHFLEX1.Col
    

    If bFlexQ Then
        If lcol = 1 And MSHFLEX1.Row > 0 Then
        
            lrow = MSHFLEX1.Row
    
            MSHFLEX1.Row = lrow
            MSHFLEX1.Col = 1
            sBdname = MSHFLEX1.Text
            cboBed.Text = sBdname
    
            Screen.MousePointer = 11
    
            ErstelleMSHFLEX
            ErstelleSQLStatement
    
            Screen.MousePointer = 0
    
            bFlexQ = False
            sBdname = ""
        End If
    End If
    
    MSHFLEX1.Col = lcol
    MSHFLEX1.sOrt = 2

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub MSHFLEX1_EnterCell()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sTT As String
    Dim lrow As Long
    Dim rs As Recordset
    Dim lLiniecol As Long
    Dim sLieferant As String
    Dim i As Integer
    
    If bFlexQ Then
        MSHFLEX1.ToolTipText = "Schauen Sie sich mit einem Doppelklick die Details an!"
    Else
        lrow = MSHFLEX1.Row
    
        MSHFLEX1.Row = 0
    
    
        If MSHFLEX1.Text = "Lieferant" Then
            MSHFLEX1.Row = lrow
            sTT = Trim(MSHFLEX1.Text)
            If sTT = "" Then
                Exit Sub
            End If
            
            If IsNumeric(sTT) = False Then
                Exit Sub
            End If
            
            sSQL = "Select Liefbez From lisrt"
            sSQL = sSQL & "  where LINR = " & sTT & ""
    
            Set rs = gdBase.OpenRecordset(sSQL)
    
            If Not rs.EOF Then
                rs.MoveFirst
                If Not IsNull(rs!LIEFBEZ) Then
                    sTT = rs!LIEFBEZ
                    MSHFLEX1.ToolTipText = sTT
                End If
            Else
                MSHFLEX1.ToolTipText = "Keine passende Bezeichnung gefunden"
            End If
            rs.Close: Set rs = Nothing
    
            
        
        ElseIf MSHFLEX1.Text = "Linie" Then
            MSHFLEX1.Row = lrow
            lLiniecol = MSHFLEX1.Col
            sTT = MSHFLEX1.Text
        
            MSHFLEX1.Row = 0
            For i = 0 To 5
            MSHFLEX1.Col = i
            
            If MSHFLEX1.Text = "Lieferant" Then
                MSHFLEX1.Row = lrow
                sLieferant = MSHFLEX1.Text
                Exit For
                
            End If
            Next i
            
            sSQL = "Select Linbezeich From linbez"
            sSQL = sSQL & "  where LPZ = " & sTT & ""
            sSQL = sSQL & " and LINR = " & sLieferant & ""
            Set rs = gdBase.OpenRecordset(sSQL)
    
            If Not rs.EOF Then
            rs.MoveFirst
                If Not IsNull(rs!LINBEZEICH) Then
                    sTT = rs!LINBEZEICH
                    MSHFLEX1.ToolTipText = sTT
                End If
            Else
                MSHFLEX1.ToolTipText = "Keine passende Bezeichnung gefunden"
            End If
            rs.Close: Set rs = Nothing
         
        ElseIf MSHFLEX1.Text = "Kunden" Then
        
            MSHFLEX1.Row = lrow
            sTT = MSHFLEX1.Text
            sSQL = "Select * From KUNDEN"
            sSQL = sSQL & "  where Kundnr = " & sTT & ""
            
            Set rs = gdBase.OpenRecordset(sSQL)
    
            If Not rs.EOF Then
            rs.MoveFirst
                If Not IsNull(rs!name) Then
                    sTT = Trim(rs!name)
                    
                End If
                If Not IsNull(rs!vorname) Then
                    sTT = sTT & "," & Space(1) & Trim(rs!vorname)
                    
                End If
                MSHFLEX1.ToolTipText = sTT
                
            Else
                MSHFLEX1.ToolTipText = "Keinen passenden Namen gefunden"
            End If
            rs.Close: Set rs = Nothing
            
        ElseIf MSHFLEX1.Text = "AGN" Then
            MSHFLEX1.Row = lrow
            sTT = MSHFLEX1.Text
            sSQL = "Select AGTEXT From agndbf"
            sSQL = sSQL & "  where AGN = " & sTT & ""
    
            Set rs = gdBase.OpenRecordset(sSQL)
    
            If Not rs.EOF Then
            rs.MoveFirst
                If Not IsNull(rs!AGTEXT) Then
                    sTT = rs!AGTEXT
                    MSHFLEX1.ToolTipText = sTT
                End If
            Else
                MSHFLEX1.ToolTipText = "Keine passende AGN-Bezeichnung gefunden"
            End If
            rs.Close: Set rs = Nothing
        Else
            'nix
            MSHFLEX1.ToolTipText = "KISS Hannover"
        End If
    End If


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_EnterCell"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub optD_Click()
    On Error GoTo LOKAL_ERROR

    Frame5.Visible = False
    lblAnzeige.Caption = ""
    lblAnzeige.Refresh
    
    framegroesse1
    Frame3.Visible = False
    lstLinA.Clear
    Frame2.Visible = False
    caldate.Visible = False
    MSHFLEX1.Visible = False
    cboSort1.SetFocus
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optD_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub optq_Click()
    On Error GoTo LOKAL_ERROR
    
    Frame5.Visible = False
    cboBed.Text = "alle"
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optq_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub optqp_Click()
    On Error GoTo LOKAL_ERROR
    
    Frame5.Visible = False
    cboBed.Text = "alle"
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optqp_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub optz_Click()
    On Error GoTo LOKAL_ERROR
    
    Frame5.Visible = True

    cboBed.Text = "alle"

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "optz_Click"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If Len(Text1(0).Text) = 0 Then
        lblAnzeige.Caption = "Geben Sie ein Anfangsdatum ein!."
        lblAnzeige.Refresh
        Text1(0).SetFocus
    End If
    
    Text1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    caldate.Visible = True
    If Index = 0 Then
        iPruef = 1
    Else
        iPruef = 2
        caldate.Value = Date
    End If
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
    MSHFLEX1.Visible = False
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Bedienerstatistik ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Programmeinstellungen auf. "
    
    Fehlermeldung1
    
End Sub

