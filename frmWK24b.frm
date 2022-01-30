VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWK24b 
   BackColor       =   &H00C0C000&
   Caption         =   "Rechnung bestimmen"
   ClientHeight    =   8610
   ClientLeft      =   2895
   ClientTop       =   2730
   ClientWidth     =   11910
   Icon            =   "frmWK24b.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8610
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Artikelpreise in Rechnung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   11
      Top             =   6600
      Width           =   7935
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   5400
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   50
         Top             =   400
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "unverändert, zzgl. MWSt"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "unverändert, ohne MWSt"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "auf Netto berechnen, zzgl. MWSt"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C000&
         Caption         =   "unverändert, inkl. MWST"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "diesen Satz drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5400
         TabIndex        =   51
         Top             =   160
         Width           =   2415
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   1
      Left            =   8880
      MaxLength       =   15
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C000&
      Caption         =   "Einstellungen Zahlungsziel"
      Height          =   735
      Left            =   9360
      TabIndex        =   24
      Top             =   8160
      Width           =   2535
      Begin VB.ComboBox cboZahlZielVoreinstellung 
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Text            =   "cboZahlZielVoreinstellung"
         Top             =   4200
         Width           =   7575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   3
         Left            =   120
         MaxLength       =   100
         TabIndex        =   35
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox txtZahlZielnach 
         Height          =   375
         Left            =   120
         MaxLength       =   200
         TabIndex        =   34
         Top             =   1920
         Width           =   7575
      End
      Begin VB.TextBox txtZahlZielvor 
         Height          =   375
         Left            =   120
         MaxLength       =   200
         TabIndex        =   32
         Top             =   480
         Width           =   7575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   120
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   """Zahlungsziel drucken"" Voreinstellung"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   45
         Top             =   3960
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inhaberangabe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   42
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblBenZeichen 
         BackStyle       =   0  'Transparent
         Caption         =   "Anzahl Zeichen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lblBenZZ 
         BackStyle       =   0  'Transparent
         Caption         =   "Anzeigetext"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   39
         Top             =   2640
         Width           =   7575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "benutzerdefinierter Text danach:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "benutzerdefinierter Text davor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zahlungsziel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "speichern"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6360
         MouseIcon       =   "frmWK24b.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   36
         ToolTipText     =   "hier alle Neuigkeiten lesen"
         Top             =   4560
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Zahlungsziel drucken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   3840
      TabIndex        =   17
      Top             =   1440
      Width           =   7935
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Der Betrag wurde Bar bezahlt."
         Height          =   255
         Index           =   16
         Left            =   3840
         TabIndex        =   49
         Top             =   3840
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Der Betrag wurde über Apple Pay bezahlt."
         Height          =   255
         Index           =   15
         Left            =   3840
         TabIndex        =   48
         Top             =   2760
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Der Betrag wurde über Google Pay bezahlt."
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   3735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Der Betrag wurde über AliPay bezahlt."
         Height          =   255
         Index           =   13
         Left            =   3840
         TabIndex        =   46
         Top             =   2400
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Betrag per Kreditkartenzahlung sofort erhalten."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   43
         Top             =   4200
         Width           =   5895
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "benutzerdefinierten Text verwenden"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   40
         Top             =   4560
         Width           =   4095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Wir bitten um Überweisung des Rechnungsbetrages innerhalb von ""[FORMEL]"" Tagen"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   7695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Bitte überweisen Sie den Betrag auf das unten stehende Konto unter Angabe Ihrer Rechnungsnummer."
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   7575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Betrag per Überweisung erhalten."
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Betrag per Sofortüberweisung erhalten."
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   28
         Top             =   3120
         Width           =   3495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Der Betrag wurde über PayPal bezahlt."
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "keine Zahlungsbedingungen drucken"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   3840
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "Der Betrag wird in den nächsten Tagen von Ihrem Konto eingezogen."
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   7335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   $"frmWK24b.frx":074C
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   7335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   $"frmWK24b.frx":07EB
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   7575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Texte und Zahlungsziel bearbeiten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3120
         MouseIcon       =   "frmWK24b.frx":0876
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   23
         ToolTipText     =   "hier alle Neuigkeiten lesen"
         Top             =   4560
         Width           =   4575
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Index           =   0
      Left            =   6360
      MaxLength       =   15
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Caption         =   "nächste Rechnungs Nr"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Listensortierung (absteigend)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   3495
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "nach Datum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C000&
         Caption         =   "nach RechnungsNr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   1
      Left            =   9720
      TabIndex        =   7
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Caption         =   "Zurück"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   7600
      TabIndex        =   6
      Top             =   7800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Caption         =   "Übernehmen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3840
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5520
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ihre Zeichen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8880
      TabIndex        =   21
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "erstellt von:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   16
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "neue Rechnungsnummer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "vorhandene Rechnungsnummern"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   9
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmWK24b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function fnPruefeDuplikatReNrWK24b() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cReNr As String
    
    cReNr = Text1.Text
    cReNr = UCase$(Trim$(cReNr))
    
    cSQL = "Select * from REKOPF where RENR = '" & cReNr & "' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        fnPruefeDuplikatReNrWK24b = 1
    Else
        fnPruefeDuplikatReNrWK24b = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeDuplikatReNrWK24b"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub LeseRechnungsnummernWK24b()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFeld As String
    Dim lWert As Long
    Dim cLBSatz As String
    
    List1.Clear
    
    cSQL = "Select RENR, REDATUM from REKOPF "
    If Option2(0).Value = True Then
        cSQL = cSQL & "order by RENR desc"
    Else
        cSQL = cSQL & "order by REDATUM desc, RENR desc"
    End If
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cLBSatz = ""
            If Not IsNull(rsrs!RENR) Then
                cFeld = rsrs!RENR
            Else
                cFeld = ""
            End If
            cFeld = Trim$(cFeld)
            cFeld = cFeld & Space$(15 - Len(cFeld))
            cLBSatz = cLBSatz & cFeld & " "
            
            If Not IsNull(rsrs!REDATUM) Then
                lWert = rsrs!REDATUM
            Else
                lWert = 0
            End If
            If lWert > 0 Then
                cFeld = Format$(lWert, "DD.MM.YYYY")
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & " "
    
            List1.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseRechnungsnummernWK24b"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim ctmp As String
    Dim cTmp2 As String
    
    Screen.MousePointer = 11
    
    If Option1(0).Value = True Then
        gcRePreisKz = "B"
    ElseIf Option1(1).Value = True Then
        gcRePreisKz = "N"
    ElseIf Option1(2).Value = True Then
        gcRePreisKz = "O"
    ElseIf Option1(3).Value = True Then
        gcRePreisKz = "Z"
    End If
    
    Select Case Index
        Case Is = 0
        
            frmWKL24!Label5(3).Caption = Text2(0).Text ' erstellt von
            frmWKL24!Label5(4).Caption = Text2(1).Text ' Ihr Zeichen
            
            If Check1(0).Value = vbChecked Then
                frmWKL24!Label5(6).Caption = Text2(4).Text ' Spezialsatz
            Else
                frmWKL24!Label5(6).Caption = "" ' Spezialsatz
            End If
            
            'Zahlungsziel
            If Option2(2).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(2).Caption
            ElseIf Option2(3).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(3).Caption
            ElseIf Option2(4).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(4).Caption
            ElseIf Option2(6).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(6).Caption
            ElseIf Option2(7).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(7).Caption
            ElseIf Option2(8).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(8).Caption
            ElseIf Option2(9).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(9).Caption
            ElseIf Option2(10).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(10).Caption
            ElseIf Option2(12).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(12).Caption
            ElseIf Option2(5).Value = True Then
                frmWKL24!Label5(5).Caption = ""
            ElseIf Option2(11).Value = True Then
                frmWKL24!Label5(5).Caption = lblBenZZ.Caption
                
                
                
            ElseIf Option2(13).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(13).Caption
            ElseIf Option2(14).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(14).Caption
            ElseIf Option2(15).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(15).Caption
            ElseIf Option2(16).Value = True Then
                frmWKL24!Label5(5).Caption = Option2(16).Caption
                
                
            End If
            
            
            gcReNr = Text1.Text
            gcReNr = Trim$(gcReNr)
            If gcReNr = "" Then
                MsgBox "Bitte eine Rechnungsnummer eingeben!", vbInformation, "Winkiss Hinweis:"
                Text1.SetFocus
            Else
                iRet = fnPruefeDuplikatReNrWK24b()
                If iRet = 0 Then
                    ctmp = "Bitte überprüfen Sie Ihre Einstellungen!" & vbCrLf & vbCrLf
                    ctmp = ctmp & "Rechnungsnummer: " & gcReNr & vbCrLf & vbCrLf
                    If gcRePreisKz = "B" Then
                        cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise enthalten MWSt-Beträge (BRUTTO)" & vbCrLf
                        cTmp2 = cTmp2 & "d.h. die im Artikelpreis enthaltene MWSt wird" & vbCrLf
                        cTmp2 = cTmp2 & "am Ende der Rechnung als 'inkl. MWSt' ausgewiesen"
                    ElseIf gcRePreisKz = "N" Then
                        cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise zzgl. MWSt. (NETTO)" & vbCrLf
                        cTmp2 = cTmp2 & "d.h. die im Artikelpreis enthaltene MWSt wird" & vbCrLf
                        cTmp2 = cTmp2 & "aus den einzelnen Positionen herausgerechnet und" & vbCrLf
                        cTmp2 = cTmp2 & "am Ende der Rechnung wieder aufaddiert ('zzgl. MWSt')"
                    ElseIf gcRePreisKz = "O" Then
                        cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise OHNE MWSt. " & vbCrLf
                        cTmp2 = cTmp2 & "d.h. es wird kein Wert für 'inkl. MWSt' " & vbCrLf
                        cTmp2 = cTmp2 & "oder 'zzgl. MWSt' ermittelt"
                    ElseIf gcRePreisKz = "Z" Then
                        cTmp2 = "Rechnungsbetrag:    " & vbCrLf & "Artikelpreise zzgl MWSt." & vbCrLf
                        cTmp2 = cTmp2 & "d.h. die Artikelpreise werden als Netto-Preise gesehen" & vbCrLf
                        cTmp2 = cTmp2 & "und für jeden Artikel wird die MWSt. berechnet, " & vbCrLf
                        cTmp2 = cTmp2 & "die am Ende der Rechnung aufaddiert wird ('zzgl. MWSt')"

                    End If
                    ctmp = ctmp & cTmp2
                    ctmp = ctmp & vbCrLf & vbCrLf
                    ctmp = ctmp & "Sind diese Angaben richtig?"
                    
                    iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Überprüfung:")
                    If iRet = vbYes Then
                        Unload frmWK24b
                    End If
                Else
                    MsgBox "Diese Rechnungsnummer ist bereits vorhanden!", vbCritical, "Winkiss Hinweis:"
                    Text1.SetFocus
                End If
            End If
            
        Case Is = 1
            
            gcReNr = ""
            Unload frmWK24b
   
        Case Is = 3
            Text1.Text = HoleNaechsteReNr
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    PositionierenWK24b
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    If NewTableSuchenDBKombi("KredVE", gdApp) Then
        If SpalteInTabellegefundenNEW("KredVE", "ZAHLZIEL", gdApp) = False Then
            SpalteAnfuegenNEW "KredVE", "ZAHLZIEL", "INTEGER", gdApp
        End If
        
        If SpalteInTabellegefundenNEW("KredVE", "ZAHLZIELVOR", gdApp) = False Then
            SpalteAnfuegenNEW "KredVE", "ZAHLZIELVOR", "TEXT(250)", gdApp
        End If
        
        If SpalteInTabellegefundenNEW("KredVE", "ZAHLZIELNACH", gdApp) = False Then
            SpalteAnfuegenNEW "KredVE", "ZAHLZIELNACH", "TEXT(250)", gdApp
        End If
        
        If SpalteInTabellegefundenNEW("KredVE", "INHABER", gdApp) = False Then
            SpalteAnfuegenNEW "KredVE", "INHABER", "TEXT(100)", gdApp
        End If
        
        If SpalteInTabellegefundenNEW("KredVE", "SpezialSatz", gdApp) = False Then
            SpalteAnfuegenNEW "KredVE", "SpezialSatz", "TEXT(250)", gdApp
        End If
    
        If SpalteInTabellegefundenNEW("KredVE", "BEZSatz", gdApp) = False Then
            SpalteAnfuegenNEW "KredVE", "BEZSatz", "INTEGER", gdApp
            sSQL = "Update kredve set bezsatz = 5"
            gdApp.Execute sSQL, dbFailOnError
        End If
    Else
        CreateTable "KREDVE", gdApp
    End If
    voreinstellungladen
    
    

    Text1.Text = ""
    Text2(0).Text = ""
    Text2(1).Text = ""
    
    LeseRechnungsnummernWK24b
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub PositionierenWK24b()
    On Error GoTo LOKAL_ERROR
    
    Frame3.Top = 1440
    Frame3.Left = 3840
    Frame3.Height = 5055
    Frame3.Width = 7935
    
    Frame4.Top = 1440
    Frame4.Left = 3840
    Frame4.Height = 5055
    Frame4.Width = 7935
    Frame4.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWK24b"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim iZahlZiel As Integer
    Dim iBezSatz As Integer
    Dim i As Integer
    Dim sTextZahlZielvor As String
    Dim sTextZahlZielnach As String
    Dim sTextSpezialSatz As String
    Dim sInhaberangabe As String
    
    loeschNEW "KREDVE", gdApp
    CreateTable "KREDVE", gdApp
    
    bo1 = Option2(0).Value
    bo2 = Option2(1).Value
    iZahlZiel = Text2(2).Text
    
    sTextSpezialSatz = Text2(4).Text
    
    sTextZahlZielvor = txtZahlZielvor.Text
    sTextZahlZielnach = txtZahlZielnach.Text
    sInhaberangabe = Text2(3).Text
    
    
    Select Case Trim(Left(cboZahlZielVoreinstellung.Text, 12))
    
        Case "Variante 1:": iBezSatz = 3
        Case "Variante 2:": iBezSatz = 2
        Case "Variante 3:": iBezSatz = 9
        Case "Variante 4:": iBezSatz = 4
        Case "Variante 5:": iBezSatz = 6
        
        Case "Variante 6:": iBezSatz = 7
        Case "Variante 7:": iBezSatz = 8
        Case "Variante 8:": iBezSatz = 10
        Case "Variante 9:": iBezSatz = 5
        Case "Variante 10:": iBezSatz = 12
        Case "Variante 11:": iBezSatz = 11
        
        Case "Variante 12:": iBezSatz = 13
        Case "Variante 13:": iBezSatz = 14
        Case "Variante 14:": iBezSatz = 15
        Case "Variante 15:": iBezSatz = 16
    
    End Select
    
    
    
    
    
    
    sSQL = "Insert into KREDVE (BO1,BO2,ZahlZiel,bezsatz,ZahlZielVor,ZahlZielNach,Inhaber,SpezialSatz) "
    sSQL = sSQL & " values (" & bo1 & "," & bo2 & "," & iZahlZiel & "," & iBezSatz & ",'" & sTextZahlZielvor & "'"
    sSQL = sSQL & ",'" & sTextZahlZielnach & "','" & sInhaberangabe & "'"
    sSQL = sSQL & " ,'" & sTextSpezialSatz & "'"
    sSQL = sSQL & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub voreinstellungladen()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim i As Integer
    For i = 2 To 16
        Option2(i).FontBold = False
        Option2(i).FontUnderline = False
    Next i
    
    Set rsrs = gdApp.OpenRecordset("KREDVE")
    Text2(2).Text = "7"
    
    If Not rsrs.EOF Then
        
        Option2(0).Value = rsrs!bo1
            
        Option2(1).Value = rsrs!bo2
        
        If Not IsNull(rsrs!INHABER) Then
            Text2(3).Text = rsrs!INHABER
        Else
            Text2(3).Text = ""
        End If
        
        If Not IsNull(rsrs!SpezialSatz) Then
            Text2(4).Text = rsrs!SpezialSatz
        Else
            Text2(4).Text = ""
        End If
        
        If Not IsNull(rsrs!ZAHLZIELNACH) Then
            txtZahlZielnach.Text = rsrs!ZAHLZIELNACH
        Else
            txtZahlZielnach.Text = ""
        End If
        
        If Not IsNull(rsrs!ZAHLZIELVOR) Then
            txtZahlZielvor.Text = rsrs!ZAHLZIELVOR
        Else
            txtZahlZielvor.Text = ""
        End If
        
        If Not IsNull(rsrs!ZAHLZIEL) Then
            Text2(2).Text = rsrs!ZAHLZIEL
        Else
            Text2(2).Text = "7"
        End If
        
        If Not IsNull(rsrs!bezsatz) Then
            Option2(rsrs!bezsatz).Value = True
            Option2(rsrs!bezsatz).FontBold = True
            Option2(rsrs!bezsatz).FontUnderline = True
        End If
            
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Dim lZahlziel As Long
    lZahlziel = Fix(Now)
    lZahlziel = lZahlziel + CInt(Text2(2).Text)
    
    Option2(3).Caption = "Bitte überweisen Sie den Betrag bis zum " & Format$(lZahlziel, "DD.MM.YYYY") & " auf das nebenstehende Konto unter Angabe Ihrer Kundennummer und Rechnungsnummer. Vielen Dank!"
    Option2(10).Caption = "Wir bitten um Überweisung des Rechnungsbetrages innerhalb von " & Text2(2).Text & " Tagen"
    
    
    cboZahlZielVoreinstellung.AddItem "Variante 1: " & Option2(3).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 2: " & Option2(2).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 3: " & Option2(9).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 4: " & Option2(4).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 5: " & Option2(6).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 6: " & Option2(7).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 7: " & Option2(8).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 8: " & Option2(10).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 9: " & Option2(5).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 10: " & Option2(12).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 11: " & Option2(11).Caption
    
    cboZahlZielVoreinstellung.AddItem "Variante 12: " & Option2(13).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 13: " & Option2(14).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 14: " & Option2(15).Caption
    cboZahlZielVoreinstellung.AddItem "Variante 15: " & Option2(16).Caption
    
    
    
    If Option2(2).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 2: " & Option2(2).Caption
    ElseIf Option2(3).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 1: " & Option2(3).Caption
    ElseIf Option2(4).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 4: " & Option2(4).Caption
    ElseIf Option2(6).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 5: " & Option2(6).Caption
    ElseIf Option2(7).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 6: " & Option2(7).Caption
    ElseIf Option2(8).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 7: " & Option2(8).Caption
    ElseIf Option2(9).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 3: " & Option2(9).Caption
    ElseIf Option2(10).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 8: " & Option2(10).Caption
    ElseIf Option2(12).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 10: " & Option2(12).Caption
    ElseIf Option2(5).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 9: " & Option2(5).Caption
    ElseIf Option2(11).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 11: " & Option2(11).Caption
        
        
        
    ElseIf Option2(13).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 12: " & Option2(13).Caption
    ElseIf Option2(14).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 13: " & Option2(14).Caption
    ElseIf Option2(15).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 14: " & Option2(15).Caption
    ElseIf Option2(16).Value = True Then
        cboZahlZielVoreinstellung.Text = "Variante 15: " & Option2(16).Caption
    End If
    
    
    
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    voreinstellungspeichern
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

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(4).ForeColor = glS1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame3_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(5).ForeColor = glS1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame4_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 4 Then
        Frame3.Visible = False
        Frame4.Visible = True
    End If
    
    If Index = 5 Then
        Frame4.Visible = False
        Frame3.Visible = True
        voreinstellungspeichern
        voreinstellungladen
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If Index = 4 Then
        Label1(4).ForeColor = glLink
    End If
    
    If Index = 5 Then
        Label1(5).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

LeseRechnungsnummernWK24b

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text1.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case 2
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text2_LostFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text2(Index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Text2(Index).BackColor = glSelBack1
    Text2(Index).SelStart = 0
    Text2(Index).SelLength = Len(Text2(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtZahlZielnach_GotFocus()
On Error GoTo LOKAL_ERROR

    txtZahlZielnach.BackColor = glSelBack1
    txtZahlZielnach.SelStart = 0
    txtZahlZielnach.SelLength = Len(txtZahlZielnach.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtZahlZielnach_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtZahlZielvor_Change()
On Error GoTo LOKAL_ERROR

    Dim lZahlziel As Long
    lZahlziel = Fix(Now)
    If Len(Text2(2).Text) > 0 Then
        lZahlziel = lZahlziel + CInt(Text2(2).Text)
    End If
    
    lblBenZZ.Caption = txtZahlZielvor.Text & " " & Format$(lZahlziel, "DD.MM.YYYY") & " " & txtZahlZielnach.Text
    lblBenZZ.Refresh
    
    lblBenZeichen.Caption = "Anzahl Zeichen: noch " & 250 - Len(lblBenZZ.Caption)
    lblBenZeichen.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtZahlZielvor_Change"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtZahlZielnach_Change()
On Error GoTo LOKAL_ERROR

    Dim lZahlziel As Long
    lZahlziel = Fix(Now)
    If Len(Text2(2).Text) > 0 Then
        lZahlziel = lZahlziel + CInt(Text2(2).Text)
    End If
    
    lblBenZZ.Caption = txtZahlZielvor.Text & " " & Format$(lZahlziel, "DD.MM.YYYY") & " " & txtZahlZielnach.Text
    lblBenZZ.Refresh
    
    lblBenZeichen.Caption = "Anzahl Zeichen: noch " & 250 - Len(lblBenZZ.Caption)
    lblBenZeichen.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtZahlZielnach_Change"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text2_Change(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 2 Then
        
        Dim lZahlziel As Long
        lZahlziel = Fix(Now)
        
        If Len(Text2(2).Text) > 0 Then
            lZahlziel = lZahlziel + CInt(Text2(2).Text)
        End If
        
        lblBenZZ.Caption = txtZahlZielvor.Text & " " & Format$(lZahlziel, "DD.MM.YYYY") & " " & txtZahlZielnach.Text
        lblBenZZ.Refresh
        
        lblBenZeichen.Caption = "Anzahl Zeichen: noch " & 250 - Len(lblBenZZ.Caption)
        lblBenZeichen.Refresh
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_Change"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub txtZahlZielvor_GotFocus()
On Error GoTo LOKAL_ERROR

    txtZahlZielvor.BackColor = glSelBack1
    txtZahlZielvor.SelStart = 0
    txtZahlZielvor.SelLength = Len(txtZahlZielvor.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtZahlZielvor_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtZahlZielnach_LostFocus()
On Error GoTo LOKAL_ERROR

    txtZahlZielnach.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtZahlZielnach_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtZahlZielvor_LostFocus()
On Error GoTo LOKAL_ERROR

    txtZahlZielvor.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtZahlZielvor_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Rechnung bestimmen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
