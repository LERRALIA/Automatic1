VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWK15a 
   BackColor       =   &H00C0C000&
   Caption         =   "WE aus Bestellung"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   405
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
   Icon            =   "frmWK15a.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cboStrichEndlos 
      Height          =   330
      Left            =   3600
      TabIndex        =   116
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00C0C000&
      Caption         =   "bestellt = geliefert"
      Height          =   255
      Left            =   120
      TabIndex        =   115
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Caption         =   "Lieferavis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5055
      Left            =   3240
      TabIndex        =   106
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   240
         TabIndex        =   107
         Top             =   960
         Width           =   6615
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   108
         Top             =   3720
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   1
         Left            =   4680
         TabIndex        =   109
         Top             =   3720
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   0
         Left            =   2450
         TabIndex        =   110
         Top             =   3720
         Width           =   2200
         _ExtentX        =   3889
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "vorhandene elektronische Lieferscheine zu folgenden Bestellungen (Auftragsnummern)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   240
         TabIndex        =   112
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00C0C000&
      Caption         =   "mit Etikett"
      Height          =   255
      Left            =   2280
      TabIndex        =   96
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   8280
      TabIndex        =   88
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C000&
      Caption         =   "Menge halten"
      Height          =   255
      Left            =   6240
      TabIndex        =   87
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   7680
      TabIndex        =   85
      Text            =   "1"
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C000&
      Caption         =   "kumulieren"
      Height          =   255
      Left            =   6240
      TabIndex        =   84
      Top             =   300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame fraArtAnfuegen 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Caption         =   "Artikel Anfügen ( Bitte mindestens ein Feld ausfüllen ) :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   9000
      TabIndex        =   44
      Top             =   8040
      Width           =   1335
      Begin VB.TextBox Text6 
         Height          =   288
         Left            =   2880
         TabIndex        =   74
         Top             =   1680
         Width           =   3732
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4545
         Left            =   240
         TabIndex        =   50
         Top             =   2280
         Width           =   11175
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   492
         Index           =   2
         Left            =   6960
         TabIndex        =   52
         Top             =   6960
         Width           =   2172
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text2 
         Height          =   288
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   49
         Top             =   600
         Width           =   1572
      End
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   2880
         TabIndex        =   48
         Top             =   960
         Width           =   2652
      End
      Begin VB.TextBox Text4 
         Height          =   288
         Left            =   2880
         TabIndex        =   47
         Top             =   1320
         Width           =   3732
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   492
         Index           =   0
         Left            =   9240
         TabIndex        =   46
         Top             =   1080
         Width           =   2172
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   " Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdAnfuegen 
         Height          =   492
         Index           =   1
         Left            =   9240
         TabIndex        =   45
         Top             =   6960
         Width           =   2172
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         TabIndex        =   51
         Top             =   2040
         Width           =   11172
      End
      Begin VB.Label label8 
         BackColor       =   &H00808000&
         Caption         =   "Lieferantenbestellnummer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   75
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel anfügen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   240
         TabIndex        =   62
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "Artikelnummer :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   55
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         Caption         =   "EAN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "Artikelbezeichnung:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   10320
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
      Begin VB.TextBox Text9 
         Height          =   255
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   117
         ToolTipText     =   "Ändern + Enter (ändert den Rechnungs-EK nur im Zugangsprotokoll, NICHT den LEK) "
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   255
         Left            =   2880
         TabIndex        =   100
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txtLieferschein 
         Height          =   255
         Left            =   120
         MaxLength       =   20
         TabIndex        =   98
         Top             =   6720
         Width           =   1455
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   4
         Left            =   9240
         TabIndex        =   95
         Top             =   6600
         Width           =   1185
         _ExtentX        =   2090
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
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "Etiketten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtStatus 
         Height          =   255
         Left            =   1560
         TabIndex        =   70
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picprogress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   3840
         ScaleHeight     =   75
         ScaleWidth      =   6915
         TabIndex        =   69
         Top             =   650
         Visible         =   0   'False
         Width           =   6975
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   3
         Left            =   5880
         TabIndex        =   65
         Top             =   6600
         Width           =   2025
         _ExtentX        =   3572
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
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "Zwischenspeichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
         Height          =   4935
         Left            =   120
         TabIndex        =   63
         Top             =   840
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   8705
         _Version        =   393216
         BackColorSel    =   10485760
         ForeColorSel    =   65535
         FocusRect       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   2
         Left            =   7920
         TabIndex        =   43
         Top             =   6600
         Width           =   1305
         _ExtentX        =   2302
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
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "Anfügen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   8640
         TabIndex        =   32
         Top             =   0
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
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
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "Geliefert + Berechnet auf Bestellt setzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   0
         Left            =   6360
         TabIndex        =   31
         Top             =   0
         Width           =   2265
         _ExtentX        =   0
         _ExtentY        =   0
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
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "Geliefert + Berechnet auf 0 setzen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Left            =   5160
         TabIndex        =   30
         Top             =   0
         Width           =   1190
         _ExtentX        =   2090
         _ExtentY        =   873
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
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Left            =   3840
         TabIndex        =   29
         Top             =   0
         Width           =   1305
         _ExtentX        =   0
         _ExtentY        =   0
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
            Name            =   "Arial"
            Size            =   8.25
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   0
         Width           =   1095
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   1
         Left            =   10440
         TabIndex        =   9
         Top             =   6600
         Width           =   1215
         _ExtentX        =   2143
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
            Name            =   "Arial"
            Size            =   8.25
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
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   8
         Top             =   6600
         Width           =   1545
         _ExtentX        =   2725
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
            Name            =   "Arial"
            Size            =   8.25
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
      Begin sevCommand3.Command Command2 
         Height          =   360
         Index           =   16
         Left            =   3840
         TabIndex        =   101
         ToolTipText     =   "Kalender"
         Top             =   6600
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
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
            Name            =   "Arial"
            Size            =   8.25
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
         ToolTip         =   "Wählen Sie hier das Datum aus."
         ToolTipTitle    =   "Kalender"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C0C000&
         Caption         =   "nur verbuchen, Restbestellung unverändert belassen"
         Height          =   255
         Left            =   4320
         TabIndex        =   119
         Top             =   6360
         Width           =   5655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Height          =   195
         Index           =   11
         Left            =   1680
         MouseIcon       =   "frmWK15a.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   118
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferscheinnummer"
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
         Index           =   2
         Left            =   120
         MouseIcon       =   "frmWK15a.frx":074C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   111
         ToolTipText     =   "mit Doppelklick zur Auswertung"
         Top             =   6480
         Width           =   1695
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
         Index           =   1
         Left            =   3000
         MouseIcon       =   "frmWK15a.frx":0A56
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   104
         ToolTipText     =   "mit Doppelklick zur Auswertung"
         Top             =   6480
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Height          =   195
         Index           =   9
         Left            =   2160
         MouseIcon       =   "frmWK15a.frx":0D60
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   103
         Top             =   6480
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "MHD:"
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
         Index           =   19
         Left            =   2280
         MouseIcon       =   "frmWK15a.frx":106A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   102
         ToolTipText     =   "mit Doppelklick zur Auswertung"
         Top             =   6720
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferantenbezeichnung"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   91
         Top             =   555
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Linr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Index           =   6
         Left            =   120
         MouseIcon       =   "frmWK15a.frx":1374
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   90
         Top             =   300
         Width           =   975
      End
      Begin VB.Label label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   77
         Top             =   5880
         Width           =   11415
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   10920
         MouseIcon       =   "frmWK15a.frx":167E
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWK15a.frx":1988
         ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem Scanpal einlesen möchten"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblUeberschrift 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   64
         Top             =   6120
         Width           =   11415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Zeilenrabatt über alles:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   9480
      TabIndex        =   67
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin sevCommand3.Command Command6 
      Height          =   300
      Left            =   11040
      TabIndex        =   66
      Top             =   240
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Suche"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame0 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   11775
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   14
         Left            =   10290
         TabIndex        =   61
         Top             =   0
         Width           =   675
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         ToolTip         =   "Runter"
         ToolTipTitle    =   "Runter"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   12
         Left            =   8880
         TabIndex        =   60
         Top             =   0
         Width           =   700
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         ToolTip         =   "Links"
         ToolTipTitle    =   "Links"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   15
         Left            =   10970
         TabIndex        =   59
         Top             =   0
         Width           =   670
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         ToolTip         =   "Rechts"
         ToolTipTitle    =   "Rechts"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   13
         Left            =   9580
         TabIndex        =   58
         Top             =   0
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         ToolTip         =   "Rauf"
         ToolTipTitle    =   "Rauf"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   11
         Left            =   8040
         TabIndex        =   22
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "C"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   10
         Left            =   7320
         TabIndex        =   21
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   ","
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   9
         Left            =   6600
         TabIndex        =   20
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "9"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   8
         Left            =   5880
         TabIndex        =   19
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "8"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   7
         Left            =   5160
         TabIndex        =   18
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "7"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   6
         Left            =   4440
         TabIndex        =   17
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "6"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   5
         Left            =   3720
         TabIndex        =   16
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "5"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   4
         Left            =   3000
         TabIndex        =   15
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "4"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   3
         Left            =   2280
         TabIndex        =   14
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "3"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   2
         Left            =   1560
         TabIndex        =   13
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   1
         Left            =   840
         TabIndex        =   12
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "1"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   720
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
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
         Caption         =   "0"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   7
         Left            =   8040
         TabIndex        =   37
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   6
         Left            =   6600
         TabIndex        =   36
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   5
         Left            =   5160
         TabIndex        =   34
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   4
         Left            =   3720
         TabIndex        =   33
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   3
         Left            =   2280
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   2
         Left            =   1560
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   1
         Left            =   840
         TabIndex        =   24
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label0 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
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
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12120
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   11535
      Begin VB.CheckBox Check5 
         Caption         =   "mit Übernahmeprotokoll"
         Height          =   255
         Left            =   9120
         TabIndex        =   105
         Top             =   6720
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "mit Differenzprotokoll"
         Height          =   255
         Left            =   9120
         TabIndex        =   94
         Top             =   6360
         Width           =   2535
      End
      Begin VB.CheckBox Check17 
         Caption         =   "keine Restbestelldateien"
         Height          =   255
         Left            =   9120
         TabIndex        =   93
         Top             =   6000
         Width           =   2535
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   7
         Left            =   9120
         TabIndex        =   92
         Top             =   3870
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Zusammenfassen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'Kein
         Caption         =   "Frame3"
         Height          =   495
         Left            =   120
         TabIndex        =   78
         Top             =   5640
         Width           =   7935
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808000&
            Caption         =   "Auftragsnr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   3
            Left            =   6360
            TabIndex        =   97
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808000&
            Caption         =   "Lieferantenbezeichnung"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   81
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808000&
            Caption         =   "Datum"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   80
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00808000&
            Caption         =   "Dateiname"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   79
            Top             =   120
            Width           =   1575
         End
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   6
         Left            =   9120
         TabIndex        =   76
         Top             =   3480
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   5
         Left            =   9120
         TabIndex        =   73
         Top             =   3090
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Importieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   4
         Left            =   9120
         TabIndex        =   72
         Top             =   2700
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Exportieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   3
         Left            =   9120
         TabIndex        =   71
         Top             =   4650
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Kundenbestellungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Linien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   252
         Index           =   3
         Left            =   9120
         TabIndex        =   56
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Artikelnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   2
         Left            =   9120
         TabIndex        =   42
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Bestellnummer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   9120
         TabIndex        =   41
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Bezeichnung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   9120
         TabIndex        =   40
         Top             =   840
         Value           =   -1  'True
         Width           =   1575
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   9120
         TabIndex        =   4
         Top             =   7200
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Height          =   5100
         Left            =   120
         MultiSelect     =   2  'Erweitert
         TabIndex        =   39
         Top             =   480
         Width           =   8895
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
         Height          =   480
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   8895
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   1
         Left            =   9120
         TabIndex        =   3
         Top             =   2310
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   0
         Left            =   9120
         TabIndex        =   2
         Top             =   1920
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   8
         Left            =   9120
         TabIndex        =   114
         Top             =   4260
         Width           =   2535
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "als Bestellvorschlag"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Anzeige"
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
         Left            =   120
         TabIndex        =   99
         Top             =   6120
         Width           =   7935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Artikelnummer / EAN"
         Height          =   255
         Index           =   3
         Left            =   9120
         TabIndex        =   83
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Wert aller Bestellungen"
         Height          =   255
         Index           =   2
         Left            =   9120
         TabIndex        =   82
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "sortiert nach"
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
         Left            =   9120
         TabIndex        =   57
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   2
         Left            =   3480
         TabIndex        =   35
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Bestellung vom ... bei:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   3000
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
      End
   End
   Begin MSComDlg.CommonDialog cdlopen 
      Left            =   11280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Datei"
      Height          =   195
      Index           =   10
      Left            =   480
      TabIndex        =   113
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "LiefBest"
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Menge"
      Height          =   255
      Index           =   4
      Left            =   7680
      TabIndex        =   86
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "Artikelnummer / EAN"
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
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
End
Attribute VB_Name = "frmWK15a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cAnfuLinr           As String
Dim gcWEdatei           As String
Dim cAnfuegenBez        As String
Dim dAnfuLEKPR          As Double
Dim bAnfuegen           As Boolean
Dim gbAender            As Boolean
Dim gbUpdate            As Boolean
Dim sZufall             As String
Dim aBreite(0 To 16)    As Integer
Dim glmaxtabzeilenanz   As Integer
Dim ws                  As Workspace
Dim cSort               As String
Dim isEtidruFree        As Boolean
Dim gbAenderKVK         As Boolean
Private Sub PositionierenWK15a()
    On Error GoTo LOKAL_ERROR
    
    With Frame0
        .Top = 7800
        .Left = 0
        .Height = 1095
        .Width = 11895
    End With

    With Frame1
        .Top = 720
        .Left = 0
        .Height = 7935
        .Width = 11895
    End With
    
    With Frame2
        .Top = 720
        .Left = 0
        .Height = 6975
        .Width = 11895
    End With
    
    With fraArtAnfuegen
        .Top = 720
        .Height = 7935
        .Left = 120
        .Width = 11655
    End With
    
    With Frame4
        .Top = 1920
        .Left = 2520
        .Height = 4455
        .Width = 7095
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ArtikelAnfuegenWKL15a()
    On Error GoTo LOKAL_ERROR

    Dim lcount      As Long
    Dim cSQL        As String
    Dim cLBSatz     As String
    Dim iTmp        As Integer
    Dim rsrs        As Recordset
    Dim rsrs1       As Recordset
    Dim rsRsF       As Recordset

    cLBSatz = ""
    For lcount = 0 To List3.ListCount - 1
        If List3.Selected(lcount) = True Then
            cLBSatz = Trim$(List3.list(lcount))
            cLBSatz = Left(cLBSatz, 6)
            cLBSatz = Trim$(cLBSatz)
            cSQL = "Select * from anfue where ARTNR = " & cLBSatz
            Set rsRsF = gdBase.OpenRecordset(cSQL)

            cSQL = "Select * from  " & sZufall & "  where ARTNR = " & cLBSatz
            Set rsrs = gdBase.OpenRecordset(cSQL)

            If rsrs.EOF Then
                rsrs.AddNew
                rsrs!artnr = cLBSatz
                rsrs!BEZEICH = rsRsF!BEZEICH
                cAnfuegenBez = rsRsF!BEZEICH
                If Not IsNull(rsRsF!lekpr) Then
                    rsrs!lekpr = rsRsF!lekpr
                    dAnfuLEKPR = Format$(rsRsF!lekpr, "#####0.00")
                Else
                    rsrs!lekpr = 0
                    dAnfuLEKPR = 0
                End If
                rsrs!BESTELLT = 0
                rsrs!GELIEFERT = 0
                rsrs!BERECHNET = 0
                rsrs!LIEFBETRAG = 0
                rsrs!ZEILEN_RAB = 0
                rsrs!ZEILENWERT = 0
                rsrs!RECHN_RAB = 0
                rsrs!RECHN_WERT = 0
                rsrs!STCK_PREIS = 0
                rsrs!linr = cAnfuLinr
                If Not IsNull(rsRsF!LIBESNR) Then
                    rsrs!LIBESNR = rsRsF!LIBESNR
                Else
                    rsrs!LIBESNR = 0
                End If
                rsrs!KVKPR1 = rsRsF!KVKPR1
                rsrs.Update
            End If
            rsrs.Close: Set rsrs = Nothing
            rsRsF.Close
            Exit For
        End If
    Next lcount

    MoveBestell2GridWK15a cSort
    fraArtAnfuegen.Visible = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ArtikelAnfuegenWKL15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AktualisiereEingangWK15a()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lRows       As Long
    Dim lartnr      As Long
    Dim lBestellt   As Long
    Dim lGeliefert  As Long
    Dim lBerechnet  As Long
    Dim dStckPreis  As Double
    Dim cFeld       As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim siAnzeige       As Single
    
    Dim lcount As Long
    Dim iStufe As Integer
   
    
    txtStatus.Text = ""
    picprogress.Visible = True
    
    
    
    
    
    Screen.MousePointer = 11
    
    MSFlexGrid1.Redraw = False
    
    lRows = MSFlexGrid1.Rows
    For lrow = 2 To lRows - 1
         
        
        siAnzeige = siAnzeige + 1
        txtStatus.Text = CStr((100 * siAnzeige) / lRows)
    
        MSFlexGrid1.Row = lrow
        
        MSFlexGrid1.Col = 0
        cFeld = MSFlexGrid1.Text
        lartnr = Val(cFeld)
        
        MSFlexGrid1.Col = 4
        cFeld = MSFlexGrid1.Text
        lBestellt = Val(cFeld)
        
        MSFlexGrid1.Col = 5
        cFeld = MSFlexGrid1.Text
        lGeliefert = Val(cFeld)
        
        MSFlexGrid1.Col = 6
        cFeld = MSFlexGrid1.Text
        lBerechnet = Val(cFeld)
        
        MSFlexGrid1.Col = 12
        cFeld = MSFlexGrid1.Text
        cFeld = fnMoveComma2Point$(cFeld)
        dStckPreis = Val(cFeld)
        
        cSQL = "Select * from  " & sZufall & "  where ARTNR = " & Trim$(Str$(lartnr))
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.Edit
            rsrs!BESTELLT = lBestellt
            rsrs!GELIEFERT = lGeliefert
            rsrs!BERECHNET = lBerechnet
            rsrs!STCK_PREIS = dStckPreis
            rsrs.Update
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
    Next lrow
    
    MSFlexGrid1.Redraw = True
    
    txtStatus.Text = ""
    picprogress.Visible = True
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AktualisiereEingangWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FormatiereGridWK15a()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    With MSFlexGrid1
        .Rows = 2
        .Cols = 17
        .FixedRows = 1
        .FixedCols = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "ArtNr."
        
        .Col = 1
        .Text = "LiefBestNr."
        
        .Col = 2
        .Text = "Bezeich"
        
        .Col = 3
        .Text = "LEkPr"
        
        .Col = 4
        .Text = "Bestellt"
        
        .Col = 5
        .Text = "Geliefert"
        
        .Col = 6
        .Text = "Berechnet"
        
        .Col = 7
        .Text = "Lieferbetrag"
        
        .Col = 8
        .Text = "Zeilenrabatt"
        
        .Col = 9
        .Text = "Zeilenwert"
        
        .Col = 10
        .Text = "Rechn.Rabatt"
        
        .Col = 11
        .Text = "Rechn.Wert"
        
        .Col = 12
        .Text = "Stückpreis"
        
        .Col = 13
        .Text = "Kassen-VK"
        
        .Col = 14
        .Text = "HLieferant"
        
        .Col = 15
        .Text = "Lieferant"
        
        .Col = 16
        .Text = "Linie"
        
        For i = 0 To 16
            aBreite(i) = TextWidth(.TextMatrix(0, i))
        Next i
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereGridWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub HoleBestellDateiWK15a(cdatei As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim cPfad       As String
    Dim cTabelle    As String
    Dim rsrs        As Recordset
    
    Dim rsArtikel   As Recordset
    Dim rsTabelle   As Recordset
    Dim cArtNr      As String
    Dim ctmp        As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cTabelle = cdatei
    
    cAnfuLinr = Left(cTabelle, Len(cTabelle) - 1)
    cAnfuLinr = Right(cAnfuLinr, Len(cAnfuLinr) - 1)
    
    Label3(6).Caption = cAnfuLinr
    Label3(7).Caption = ermLiefBez(CLng(cAnfuLinr))

    Set rsTabelle = gdBase.OpenRecordset(cTabelle, dbOpenTable)
    
    If Not rsTabelle.EOF Then
        rsTabelle.MoveFirst
        Do While Not rsTabelle.EOF
            If Not IsNull(rsTabelle!artnr) Then
                cArtNr = Trim(rsTabelle!artnr)
            End If
            cSQL = "Select * from Artikel where artnr = " & cArtNr
            Set rsArtikel = gdBase.OpenRecordset(cSQL)
            If Not rsArtikel.EOF Then
                rsArtikel.MoveFirst
                
                rsTabelle.Edit
                If Not IsNull(rsArtikel!KVKPR1) Then
                    rsTabelle!KVKPR1 = rsArtikel!KVKPR1
                End If
                
                If Not IsNull(rsArtikel!vkpr) Then
                    rsTabelle!vkpr = rsArtikel!vkpr
                End If
                rsTabelle.Update
            End If
            
            cSQL = "Select * from Artlief where artnr = " & cArtNr
            cSQL = cSQL & " and Linr = " & cAnfuLinr
            Set rsArtikel = gdBase.OpenRecordset(cSQL)
            If Not rsArtikel.EOF Then
                rsArtikel.MoveFirst
                rsTabelle.Edit
                If Not IsNull(rsArtikel!lekpr) Then
                    rsTabelle!lekpr = rsArtikel!lekpr
                End If
                rsTabelle.Update
            End If
            
            rsArtikel.Close
        rsTabelle.MoveNext
        Loop
    End If
    rsTabelle.Close
    
    
    Dim cMitt As String
    Label8(2).Caption = ""
    cSQL = "Select distinct(Mitteilung) as mitt from " & cTabelle
    Set rsTabelle = gdBase.OpenRecordset(cSQL)
    
    If Not rsTabelle.EOF Then
        If Not IsNull(rsTabelle!mitt) Then
            
            cMitt = Trim(rsTabelle!mitt)
            cMitt = SwapStr(cMitt, Chr(13), " ")
            cMitt = SwapStr(cMitt, Chr(10), " ")
            Label8(2).Caption = cMitt
        End If
    End If
    rsTabelle.Close
    
    loesch sZufall
    
    cSQL = "Create Table " & sZufall
    cSQL = cSQL & " ( ARTNR Long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", LEKPR Double"
    cSQL = cSQL & ", BESTELLT Long"
    cSQL = cSQL & ", GELIEFERT Long"
    cSQL = cSQL & ", BERECHNET Long"
    cSQL = cSQL & ", LIEFBETRAG Double"
    cSQL = cSQL & ", ZEILEN_RAB Double"
    cSQL = cSQL & ", ZEILENWERT Double"
    cSQL = cSQL & ", RECHN_RAB Double"
    cSQL = cSQL & ", RECHN_WERT Double"
    cSQL = cSQL & ", STCK_PREIS Double"
    cSQL = cSQL & ", LINR Long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", KVKPR1 Double"
    cSQL = cSQL & ", LPZ Long"
    cSQL = cSQL & ", MOPREIS Long"
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index ARTNR on " & sZufall & " (ARTNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update " & cTabelle & " set Libesnr = '' where LIBESNR is null"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into " & sZufall
    cSQL = cSQL & " Select ARTNR, BEZEICH, LEKPR, BESTVOR as BESTELLT"
    cSQL = cSQL & ", BESTVOR as GELIEFERT, BESTVOR as BERECHNET"
    cSQL = cSQL & ", LEKPR * BESTVOR as LIEFBETRAG, 0 as ZEILEN_RAB"
    cSQL = cSQL & ", LEKPR * BESTVOR as ZEILENWERT, 0 as RECHN_RAB"
    cSQL = cSQL & ", LEKPR * BESTVOR as RECHN_WERT, LEKPR as STCK_PREIS"
    cSQL = cSQL & ", LINR, LIBESNR, KVKPR1,LPZ ,mopreis "
    cSQL = cSQL & " from " & cTabelle & " where BESTVOR <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    If Datendrin(sZufall, gdBase) Then
        FormatiereGridWK15a
        
        If Option1(0).Value = True Then
            cSort = " order by MOPREIS, BEZEICH"
        ElseIf Option1(1).Value = True Then
            cSort = " order by MOPREIS, val(LIBESNR) asc "
        ElseIf Option1(2).Value = True Then
            cSort = " order by MOPREIS, ARTNR"
        ElseIf Option1(3).Value = True Then
            cSort = " order by MOPREIS, LPZ"
        Else
            cSort = " order by MOPREIS, LPZ"
        End If
        
        MoveBestell2GridWK15a cSort
    
        Frame2.Visible = True
        Frame0.Visible = True
        Frame1.Enabled = False
    Else
        cdatei = List2.list(List2.ListIndex)
        cdatei = Trim(Left(cdatei, 10))
        cdatei = UCase$(cdatei)
    
        cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW cdatei, gdBase
        LeseInhaltWK15a
    End If
    
Exit Sub

LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "HoleBestellDateiWK15a"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub check_Budni_Lieferavis(cdatei As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sCheckLinr  As String
    Dim sSQL        As String
    
    sCheckLinr = Left(cdatei, Len(cdatei) - 1)
    sCheckLinr = Right(sCheckLinr, Len(sCheckLinr) - 1)
                
    Dim sBudniKundnr As String
    Dim rsLi As DAO.Recordset
    
    sBudniKundnr = ""
    
    sSQL = "select KUNDNR from LISRT where FORMAT = 'EDIBUDNI' and Linr = " & sCheckLinr
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        sBudniKundnr = Trim(rsLi!Kundnr)
    End If
    rsLi.Close: Set rsLi = Nothing
    
    If Val(sBudniKundnr) > 0 Then

        'dann bau mal die Verbindung auf und hol alles für den Kunden ab

        giKissFtpMode = 36
        frmWKL38.Show 1

    End If
    
    'vorhandene Budni-lieferavis csv in Auftragstabellen umwandeln
    in_Kiss_Lieferavis_wandeln "BUDNI", sCheckLinr
    
Exit Sub

LOKAL_ERROR:
    
     Fehler.gsDescr = err.Description
     Fehler.gsNumber = err.Number
     Fehler.gsFormular = Me.name
     Fehler.gsFunktion = "check_Budni_Lieferavis"
     Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
     
     Fehlermeldung1

End Sub


Private Sub Create_Bestell_Datei(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim cPfad       As String
    Dim cTabelle    As String
    Dim rsrs        As Recordset
    
    Dim rsArtikel   As Recordset
    Dim rsTabelle   As Recordset
    Dim cArtNr      As String
    Dim ctmp        As String
    
    Dim iZufall     As Integer
    
    Randomize
    iZufall = 0
    iZufall = Int((999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 999 generieren.
    sZufall = Str(iZufall)
    sZufall = "A" & Trim(sZufall)
    
    gcWEdatei = "BESTIMP"
    cAnfuLinr = CStr(lLinr)
    
    
    Label3(6).Caption = CStr(lLinr)
    Label3(7).Caption = ermLiefBez(lLinr)

    
    loeschNEW sZufall, gdBase
    
    cSQL = "Create Table " & sZufall
    cSQL = cSQL & " ( ARTNR Long"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", LEKPR Double"
    cSQL = cSQL & ", BESTELLT Long"
    cSQL = cSQL & ", GELIEFERT Long"
    cSQL = cSQL & ", BERECHNET Long"
    cSQL = cSQL & ", LIEFBETRAG Double"
    cSQL = cSQL & ", ZEILEN_RAB Double"
    cSQL = cSQL & ", ZEILENWERT Double"
    cSQL = cSQL & ", RECHN_RAB Double"
    cSQL = cSQL & ", RECHN_WERT Double"
    cSQL = cSQL & ", STCK_PREIS Double"
    cSQL = cSQL & ", LINR Long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", KVKPR1 Double"
    cSQL = cSQL & ", LPZ Long"
    cSQL = cSQL & ", MOPREIS Long"
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index ARTNR on " & sZufall & " (ARTNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Insert into " & sZufall
    cSQL = cSQL & " Select ARTNR, Menge as BESTELLT"
    cSQL = cSQL & " ,Menge as GELIEFERT, Menge as BERECHNET "
    cSQL = cSQL & " , " & lLinr & " as linr "
    cSQL = cSQL & " from BESTIMP where Menge <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update " & sZufall & " Z inner join Artikel A on Z.Artnr = A.artnr"
    cSQL = cSQL & " SET Z.BEZEICH = A.BEZEICH "
    cSQL = cSQL & " , Z.KVKPR1 = A.KVKPR1 "
    cSQL = cSQL & " , Z.LPZ = A.LPZ "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update " & sZufall & " Z inner join ARTLIEF A on Z.Artnr = A.artnr and Z.linr = A.linr"
    cSQL = cSQL & " SET Z.LEKPR = A.LEKPR "
    cSQL = cSQL & " , Z.LIBESNR = A.LIBESNR "
    gdBase.Execute cSQL, dbFailOnError
       
    cSQL = "Update " & sZufall & " set "
    cSQL = cSQL & " LIEFBETRAG = LEKPR * BESTELLT "
    cSQL = cSQL & ", ZEILEN_RAB = 0 "
    cSQL = cSQL & ", RECHN_RAB = 0 "
    cSQL = cSQL & ", ZEILENWERT = LEKPR * BESTELLT "
    cSQL = cSQL & ", RECHN_WERT = LEKPR * BESTELLT "
    cSQL = cSQL & ", STCK_PREIS = LEKPR "
    gdBase.Execute cSQL, dbFailOnError
    
    If Datendrin(sZufall, gdBase) Then
        FormatiereGridWK15a
        
        If Option1(0).Value = True Then
            cSort = " order by MOPREIS, BEZEICH"
        ElseIf Option1(1).Value = True Then
            cSort = " order by MOPREIS, val(LIBESNR) asc "
        ElseIf Option1(2).Value = True Then
            cSort = " order by MOPREIS, ARTNR"
        ElseIf Option1(3).Value = True Then
            cSort = " order by MOPREIS, LPZ"
        Else
            cSort = " order by MOPREIS, LPZ"
        End If
        
        MoveBestell2GridWK15a cSort
    
        Frame2.Visible = True
        Frame0.Visible = True
        Frame1.Enabled = False
    
    End If
    
Exit Sub

LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Create_Bestell_Datei"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Function ermLastfocus(sTabname) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermLastfocus = ""
    
    cSQL = "Select ARTNR from LASTFOCUS where TABNAME like '" & sTabname & "*' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!artnr) Then
            ermLastfocus = rsrs!artnr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermLastfocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermLastSortierung(sTabname) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermLastSortierung = ""
    
    cSQL = "Select Sortierung from LASTFOCUS where TABNAME like '" & sTabname & "*' "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Sortierung) Then
            ermLastSortierung = rsrs!Sortierung
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermLastSortierung"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub HoleZwischengespeichertebestelldatei(cdatei As String)
    On Error GoTo LOKAL_ERROR
    
    Dim rsQZW       As Recordset
    Dim i           As Integer
    Dim j           As Integer
    Dim lrow        As Long
    Dim lRows       As Long
    Dim ctmp        As String
    Dim cPfad       As String
    Dim sSQL        As String
    Dim cTabelle    As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cTabelle = cdatei
    
    cAnfuLinr = Left(cTabelle, Len(cTabelle) - 1)
    cAnfuLinr = Right(cAnfuLinr, Len(cAnfuLinr) - 1)
    
    Label3(6).Caption = cAnfuLinr
    Label3(7).Caption = ermLiefBez(CLng(cAnfuLinr))
        
    loesch sZufall

    sSQL = "Create Table " & sZufall
    sSQL = sSQL & " ( ARTNR Long"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LEKPR Double"
    sSQL = sSQL & ", BESTELLT Long"
    sSQL = sSQL & ", GELIEFERT Long"
    sSQL = sSQL & ", BERECHNET Long"
    sSQL = sSQL & ", LIEFBETRAG Double"
    sSQL = sSQL & ", ZEILEN_RAB Double"
    sSQL = sSQL & ", ZEILENWERT Double"
    sSQL = sSQL & ", RECHN_RAB Double"
    sSQL = sSQL & ", RECHN_WERT Double"
    sSQL = sSQL & ", STCK_PREIS Double"
    sSQL = sSQL & ", LINR Long"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", KVKPR1 Double"
    sSQL = sSQL & ", LPZ Long"
    sSQL = sSQL & ", MOPREIS Long"
    sSQL = sSQL & ")"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Create Index ARTNR on " & sZufall & " (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from " & cTabelle & " where artnr not in (Select artnr from artlief where linr = " & cAnfuLinr & ")"
    gdBase.Execute sSQL, dbFailOnError
    

    sSQL = "Insert into " & sZufall
    sSQL = sSQL & " Select ARTNR "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", LEKPR "
    sSQL = sSQL & ", BESTELLT "
    sSQL = sSQL & ", GELIEFERT "
    sSQL = sSQL & ", BERECHNET "
    sSQL = sSQL & ", lief as liefbetrag "
    sSQL = sSQL & ", zeilenrab as ZEILEN_RAB "
    sSQL = sSQL & ", zeile as zeilenwert "
    sSQL = sSQL & ", rechnrab as RECHN_RAB"
    sSQL = sSQL & ", rechn as RECHN_WERT "
    sSQL = sSQL & ", stck as STCK_PREIS "
    sSQL = sSQL & ", LINR "
    sSQL = sSQL & ", KVKPR1 "
    sSQL = sSQL & ", LIBESNR "
    sSQL = sSQL & ", LPZ "
    sSQL = sSQL & ", Mopreis "
    sSQL = sSQL & " from " & cTabelle
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    FormatiereGridWK15a
    
    MoveBestell2GridWK15a cSort
    
    Frame2.Visible = True
    Frame0.Visible = True
    Frame1.Enabled = False
    
Exit Sub

LOKAL_ERROR:
    
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 3008 Or err.Number = 3218 Then 'Datenbanksperrung
        giErrorZaehler = giErrorZaehler + 1
        
        If giErrorZaehler > 2 Then

            ctmp = "Es wurde jetzt " & giErrorZaehler & " Mal versucht diese Aktion durchzuführen - leider erfolglos."
            ctmp = ctmp & " Drücken Sie auf ' Wiederholen ' oder probieren Sie es zu einem späteren Zeitpunkt."
            
            giErrorZaehler = 0
            
            dlgAbfrage.BCaptioneins = "Wiederholen"
            dlgAbfrage.BCaptionzwei = "Abbrechen"
            dlgAbfrage.Überschrift = "Datenbank Hinweis:"
            dlgAbfrage.Beschriftung = ctmp
            dlgAbfrage.Show vbModal
            
            If dlgAbfrage.Back = 1 Then
                Command1_Click 0
            Else
                Exit Sub
            End If
        
        Else
            Command1_Click 0 'nochmal
        End If
    
    Else
    

        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "HoleZwischengespeichertebestelldatei"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

        Fehlermeldung1
        Resume Next

    End If
End Sub
Private Sub LeseInhaltWK15a()
    On Error GoTo LOKAL_ERROR
        
    List1.Clear
    List2.Clear
    List1.AddItem "Dateiname               Bestellinformationen                Auftragswert AuftragNr"
    
    If Option2(0).Value = True Then
        ListeFuellAnfangsbuchdataT "Q", List2, "tabname", Label3(3)
    ElseIf Option2(1).Value = True Then
        ListeFuellAnfangsbuchdataT "Q", List2, "tabdate", Label3(3)
    ElseIf Option2(2).Value = True Then
        ListeFuellAnfangsbuchdataT "Q", List2, "Liefbez", Label3(3)
    ElseIf Option2(3).Value = True Then
        ListeFuellAnfangsbuchdataT "Q", List2, "AUFTRAGSNR", Label3(3)
    End If
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseInhaltWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MoveBestell2GridWK15a(cSort As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lSummeBestellt      As Long
    Dim lSummeGeliefert     As Long
    Dim lSummeBerechnet     As Long
    Dim lRows               As Long
    Dim lrow                As Long
    Dim lcol                As Long
    Dim lPos                As Long
    Dim ctmp                As String
    Dim cSQL                As String
    Dim sWert               As String
    Dim dLiefWert           As Double
    Dim dEkpr               As Double
    Dim dBestellt           As Double
    Dim dBerechnet          As Double
    Dim dGeliefert          As Double
    Dim dZeilenRabatt       As Double
    Dim dZeilenWert         As Double
    Dim dRechnRabatt        As Double
    Dim dRechnWert          As Double
    Dim dWert               As Double
    Dim dSummeLiefWert      As Double
    Dim dSummeZeilenWert    As Double
    Dim dSummeRechWert      As Double
    Dim rsrs                As Recordset
    Dim i                   As Integer
    Dim j                   As Integer
    
    cSQL = "Select * from " & sZufall & cSort
    
    MSFlexGrid1.Redraw = False
    
    glmaxtabzeilenanz = 0
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lRows = rsrs.RecordCount
        glmaxtabzeilenanz = lRows
        MSFlexGrid1.Rows = lRows + 2
        rsrs.MoveFirst
        
        lrow = 1
        lSummeBestellt = 0
        lSummeGeliefert = 0
        lSummeBerechnet = 0
        dSummeLiefWert = 0
        dSummeZeilenWert = 0
        dSummeRechWert = 0
        Do While Not rsrs.EOF
            lrow = lrow + 1
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.Col = 0
            If Not IsNull(rsrs!artnr) Then
                MSFlexGrid1.Text = rsrs!artnr
            Else
                MSFlexGrid1.Text = "-1"
            End If
            If lrow = 1 Then
                Label0(2).Caption = MSFlexGrid1.Text
            End If
            MSFlexGrid1.Col = 1
            
            If Not IsNull(rsrs!LIBESNR) Then
            
                sWert = rsrs!LIBESNR
'                MSFlexGrid1.Text = " " & rsrs!LIBESNR
            Else
                sWert = ""
'                MSFlexGrid1.Text = ""
            End If
            
            MSFlexGrid1.Text = sWert
            
            MSFlexGrid1.Col = 2
            If Not IsNull(rsrs!BEZEICH) Then
                ctmp = rsrs!BEZEICH
            Else
                ctmp = ""
            End If
            Do
                lPos = InStr(1, ctmp, "  ")
                If lPos <> 0 Then
                    ctmp = Left(ctmp, lPos - 1) & Right(ctmp, Len(ctmp) - lPos)
                End If
            Loop While lPos <> 0
            
            MSFlexGrid1.Text = ctmp
            
            MSFlexGrid1.Col = 3
            If Not IsNull(rsrs!lekpr) Then
                dWert = rsrs!lekpr
            Else
                dWert = 0
            End If
            dEkpr = dWert
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
                    
            MSFlexGrid1.Col = 4
            If Not IsNull(rsrs!BESTELLT) Then
                dWert = rsrs!BESTELLT
            Else
                dWert = 0
            End If
            dBestellt = dWert
            lSummeBestellt = lSummeBestellt + dWert
            MSFlexGrid1.Text = Format$(dWert, "#####0")
                    
            MSFlexGrid1.Col = 5
            If Not IsNull(rsrs!GELIEFERT) Then
                dWert = rsrs!GELIEFERT
            Else
                dWert = 0
            End If
            dGeliefert = dWert
            lSummeGeliefert = lSummeGeliefert + dWert
            MSFlexGrid1.Text = Format$(dWert, "#####0")
                    
            MSFlexGrid1.Col = 6
            If Not IsNull(rsrs!BERECHNET) Then
                dWert = rsrs!BERECHNET
            Else
                dWert = 0
            End If
            dBerechnet = dWert
            lSummeBerechnet = lSummeBerechnet + dWert
            MSFlexGrid1.Text = Format$(dWert, "#####0")
                    
            MSFlexGrid1.Col = 7
            dLiefWert = dEkpr * dBerechnet
            
            dSummeLiefWert = dSummeLiefWert + dLiefWert
            MSFlexGrid1.Text = Format$(dLiefWert, "#####0.00")
                    
            MSFlexGrid1.Col = 8
            If Not IsNull(rsrs!ZEILEN_RAB) Then
                dWert = rsrs!ZEILEN_RAB
            Else
                dWert = 0
            End If
            dZeilenRabatt = dWert
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
                    
            MSFlexGrid1.Col = 9
            If dZeilenRabatt > 0 Then
                dZeilenWert = dLiefWert * ((100 - dZeilenRabatt) / 100)
            Else
                dZeilenWert = dLiefWert
            End If
            dSummeZeilenWert = dSummeZeilenWert + dZeilenWert
            MSFlexGrid1.Text = Format$(dZeilenWert, "#####0.00")
                    
            MSFlexGrid1.Col = 10
            If Not IsNull(rsrs!RECHN_RAB) Then
                dWert = rsrs!RECHN_RAB
            Else
                dWert = 0
            End If
            dRechnRabatt = dWert
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
                    
            MSFlexGrid1.Col = 11
            If dRechnRabatt > 0 Then
                dRechnWert = dZeilenWert * ((100 - dRechnRabatt) / 100)
            Else
                dRechnWert = dZeilenWert
            End If
            dSummeRechWert = dSummeRechWert + dRechnWert
            MSFlexGrid1.Text = Format$(dRechnWert, "#####0.00")
                    
            MSFlexGrid1.Col = 12
            If dGeliefert <> 0 Then
                dWert = dRechnWert / dGeliefert
            Else
                dWert = 0
            End If
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
                    
            MSFlexGrid1.Col = 13
            If Not IsNull(rsrs!KVKPR1) Then
                dWert = rsrs!KVKPR1
            Else
                dWert = 0
            End If
            MSFlexGrid1.Text = Format$(dWert, "#####0.00")
            
            
            MSFlexGrid1.Col = 14
            If Not IsNull(rsrs!MOPREIS) Then
                dWert = rsrs!MOPREIS
            Else
                dWert = 0
            End If
            MSFlexGrid1.Text = dWert
            
            MSFlexGrid1.Col = 15
            If Not IsNull(rsrs!linr) Then
                dWert = rsrs!linr
            Else
                dWert = 0
            End If
            MSFlexGrid1.Text = dWert
            
            
            MSFlexGrid1.Col = 16
            If Not IsNull(rsrs!LPZ) Then
                dWert = rsrs!LPZ
            Else
                dWert = 0
            End If
            MSFlexGrid1.Text = dWert
            
            
            If dGeliefert < dBestellt Then
                For i = 6 To 14
                    MSFlexGrid1.Col = i
                    MSFlexGrid1.CellBackColor = vbRed
                Next i
            Else
                For i = 6 To 14
                    MSFlexGrid1.Col = i
                    MSFlexGrid1.CellBackColor = vbWhite
                Next i

            End If

            
            
            rsrs.MoveNext
        Loop
        
        
        ctmp = "Bestellt: "
        ctmp = ctmp & Format$(lSummeBestellt, "#####0")
        ctmp = ctmp & " Geliefert: "
        ctmp = ctmp & Format$(lSummeGeliefert, "#####0")
        ctmp = ctmp & " Berechnet: "
        ctmp = ctmp & Format$(lSummeBerechnet, "#####0")
        ctmp = ctmp & " Lieferwert: "
        ctmp = ctmp & Format$(dSummeLiefWert, "#####0.00")
        ctmp = ctmp & " Zeilenwert: "
        ctmp = ctmp & Format$(dSummeZeilenWert, "#####0.00")
        ctmp = ctmp & " Rechnungswert: "
        ctmp = ctmp & Format$(dSummeRechWert, "#####0.00")
        
        lblUeberschrift(1).Caption = ctmp
        
        MSFlexGrid1.RowHeight(1) = 0
        
        With MSFlexGrid1
            For i = 0 To .Rows - 1
                .Row = i
                For j = 0 To .Cols - 1
                .Col = j
                    If TextWidth(.TextMatrix(i, j)) > aBreite(j) Then
                        aBreite(j) = TextWidth(.TextMatrix(i, j))
                    End If
                Next j
            Next i
            
            For i = 0 To .Cols - 1
                .Col = i
                .ColWidth(i) = aBreite(i) * 1.7
            Next i
        End With
        
        MSFlexGrid1.Redraw = True
        
        Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
        
        If MSFlexGrid1.Visible = False Then
            MSFlexGrid1.Row = 1
            MSFlexGrid1.Col = 3
            Label0(0).Caption = "1"
            Label0(1).Caption = "3"
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveBestell2GridWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
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
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check17_Click()
On Error GoTo LOKAL_ERROR

    speicherLNull
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check17_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherLNull()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If Check17.Value = vbChecked Then
        sSQL = "Update DBEINSTE Set DELBDAT = true "
        gdBase.Execute sSQL, dbFailOnError

        gbDELBDAT = True
    Else
        sSQL = "Update DBEINSTE Set DELBDAT = false "
        gdBase.Execute sSQL, dbFailOnError

        gbDELBDAT = False
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherLNull"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub speicherDiffProt()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If Check3.Value = vbChecked Then
        sSQL = "Update DBEINSTE Set DIFFPROT = true "
        gdBase.Execute sSQL, dbFailOnError

        gbDIFFPROT = True
    Else
        sSQL = "Update DBEINSTE Set DIFFPROT = false "
        gdBase.Execute sSQL, dbFailOnError

        gbDIFFPROT = False
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherDiffProt"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub speicherUEBERPROT()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    If Check5.Value = vbChecked Then
        sSQL = "Update DBEINSTE Set UEBERPROT = true "
        gdBase.Execute sSQL, dbFailOnError

        gbUEBERPROT = True
    Else
        sSQL = "Update DBEINSTE Set UEBERPROT = false "
        gdBase.Execute sSQL, dbFailOnError

        gbUEBERPROT = False
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherUEBERPROT"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Check3_Click()
On Error GoTo LOKAL_ERROR

    speicherDiffProt
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check5_Click()
On Error GoTo LOKAL_ERROR

    speicherUEBERPROT
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check5_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check6_Click()
On Error GoTo LOKAL_ERROR
    
    If Check6.Value = vbChecked Then
        Check2.Visible = False
    Else
        Check2.Visible = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check6_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Check2_Click()
On Error GoTo LOKAL_ERROR
    
    If Check2.Value = vbChecked Then
        Check6.Visible = False
    Else
        Check6.Visible = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub cmdAnfuegen_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0
            bAnfuegen = False
            artikel_suchen

        Case Is = 1
            fraArtAnfuegen.Visible = False
        Case Is = 2
            bAnfuegen = True
            ArtikelAnfuegenWKL15a
    End Select
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdAnfuegen_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Private Sub artikel_suchen()
    On Error GoTo LOKAL_ERROR:
    
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim sFeld       As String
    Dim rec         As Recordset
    Dim cBezeichSuche As String
    Dim cZiel           As String
    
    cZiel = gcDBPfad
    If Right$(cZiel, 1) <> "\" Then
        cZiel = cZiel & "\"
    End If
    
    loeschNEW "anfue", gdBase

    sSQL = "select artikel.artnr"
    sSQL = sSQL & ", artikel.bezeich "
    sSQL = sSQL & ", artikel.ean "
    sSQL = sSQL & ", artlief.libesnr "
    sSQL = sSQL & ", artikel.kvkpr1 "
    sSQL = sSQL & ", artlief.lekpr "
    sSQL = sSQL & " into anfue from artikel inner join artlief on "
    sSQL = sSQL & " artikel.artnr = artlief.artnr "
    sSQL = sSQL & " Where artlief.linr = " & cAnfuLinr
    sSQL = sSQL & " and artikel.artnr not in(select artnr from " & gcWEdatei & ")"
    
    If Trim(Text2.Text) <> "" Then
        sSQL = sSQL & " and artikel.artnr= " & Trim(Text2.Text)
    End If
    
    If Text4.Text <> "" Then
        cBezeichSuche = Text4.Text
        cBezeichSuche = SwapStr(cBezeichSuche, " ", "*")
        sSQL = sSQL & " and bezeich like '*" & cBezeichSuche & "*'"
    End If
    
    

    
    If Text3.Text <> "" Then
    
    
    
        If Ist_in_ARTEAN_K(Text3.Text) Then
                
        End If
        
        
        sSQL = sSQL & " and (ean like '" & Text3.Text & "*'"
        sSQL = sSQL & " or EAN2 like '" & Text3.Text & "*' "
        sSQL = sSQL & " or EAN3 like '" & Text3.Text & "*' )"
    
    
    
    
'        sSQL = sSQL & " and ean like '" & Text3.Text & "*'"
    End If
    
    If Text6.Text <> "" Then
        sSQL = sSQL & " and artlief.libesnr like '" & Text6.Text & "*'"
    End If
        
    gdBase.Execute sSQL, dbFailOnError
    
    
    artikel_zeigen "Bezeich"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikel_suchen"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub artikel_zeigen(corder As String)
    On Error GoTo LOKAL_ERROR:
    
    Dim sSQL        As String
    Dim cLBSatz     As String
    Dim sFeld       As String
    Dim rec         As Recordset
    
    List4.Clear
    List4.AddItem "ArtNr" & Space(5) & "Artikelbezeichnung" & Space(21) & "EAN" & Space(14) & "BestellNr"
    List3.Clear
    
    sSQL = " Select * from anfue order by  " & corder
    Set rec = gdBase.OpenRecordset(sSQL)
    If Not rec.EOF Then
        rec.MoveFirst
        Do While Not rec.EOF
        
           If Not IsNull(rec!artnr) Then
               sFeld = rec!artnr
           End If
           
           sFeld = sFeld & Space$(10 - Len(sFeld))
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           If Not IsNull(rec!BEZEICH) Then
               If Len(rec!BEZEICH) > 35 Then
                   sFeld = Mid$(rec!BEZEICH, 1, 32) & "..."
               Else
                   sFeld = rec!BEZEICH
               End If
           End If
           
           sFeld = sFeld & Space$(37 - Len(sFeld))
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           If Not IsNull(rec!EAN) Then
               sFeld = rec!EAN
           Else
               sFeld = ""
           End If
           
           sFeld = Space$(15 - Len(sFeld)) & sFeld
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           If Not IsNull(rec!LIBESNR) Then
               sFeld = rec!LIBESNR
           End If
           
           sFeld = Space$(13 - Len(sFeld)) & sFeld
           cLBSatz = cLBSatz & sFeld
           sFeld = ""
           
           List3.AddItem cLBSatz
           cLBSatz = ""
           rec.MoveNext
        Loop
    End If
    rec.Close: Set rec = Nothing
    
    
    List4.Refresh
    List3.Refresh
    List3.Visible = True
    List4.Visible = True
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikel_zeigen"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iWert As Integer
    Dim ctmp As String
    Dim lrow As Long
    Dim lcol As Long
    Dim iStufe As Integer
    
    If Val(Label0(1).Caption) <> 7 _
    And Val(Label0(1).Caption) <> 9 _
    And Val(Label0(1).Caption) <> 11 _
    And Val(Label0(1).Caption) <> 12 Then
    
        Select Case Index
            
            Case 0 To 9     'Ziffern 0 bis 9
                iStufe = 1
                If Label0(0).Caption <> "-999" Then
                    iStufe = 2
                    If Label0(2).Caption <> "SUMME:" Then
                        iStufe = 3
                        lrow = Val(Label0(0).Caption)
                        iStufe = 4
                        lcol = Val(Label0(1).Caption)
                        iStufe = 5
                        If giErsetzen > 0 Then
                            iStufe = 6
                            ctmp = MSFlexGrid1.Text
                            iStufe = 7
                            ctmp = ctmp & Command0(Index).Caption
                            iStufe = 8
                            MSFlexGrid1.TextMatrix(lrow, lcol) = ctmp
                            iStufe = 9
                        Else
                            MSFlexGrid1.TextMatrix(lrow, lcol) = Command0(Index).Caption
                            iStufe = 10
                        End If
                        giErsetzen = 2
                        iStufe = 11
                        gbAender = True
                        gbUpdate = True
                    End If
                Else
                    iStufe = 12
                    Text1.Text = Text1.Text & Command0(Index).Caption
                End If
                
            Case Is = 10    'Komma
                If Label0(0).Caption <> "-999" Then
                    If Label0(2).Caption <> "SUMME:" Then
                        lrow = Val(Label0(0).Caption)
                        lcol = Val(Label0(1).Caption)
                        ctmp = MSFlexGrid1.Text
                        If InStr(ctmp, ",") = 0 Then
                            If giErsetzen > 0 Then
                                ctmp = MSFlexGrid1.Text
                                ctmp = ctmp & Command0(Index).Caption
                                MSFlexGrid1.TextMatrix(lrow, lcol) = ctmp
                            Else
                                MSFlexGrid1.TextMatrix(lrow, lcol) = Command0(Index).Caption
                            End If
                            giErsetzen = 2
                            gbAender = True
                            gbUpdate = True
                        End If
                    
                    End If
                Else
                    If InStr(Text1.Text, ",") = 0 Then
                        Text1.Text = Text1.Text & Command0(Index).Caption
                    End If
                End If
            Case Is = 11    'Clear
                If Label0(0).Caption <> "-999" Then
                    If Label0(2).Caption <> "SUMME:" Then
                        lrow = Val(Label0(0).Caption)
                        lcol = Val(Label0(1).Caption)
                        MSFlexGrid1.TextMatrix(lrow, lcol) = ""
                        gbAender = True
                        gbUpdate = True
                    End If
                Else
                    Text1.Text = ""
                End If
            
            Case Is = 12    'Links
                If Label0(0).Caption <> "-999" Then
                    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
                    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
                    iWert = Val(Label0(1).Caption)
                    If iWert > 3 Then
                        iWert = iWert - 1
                    End If
                    Label0(4).Caption = Trim$(Str$(MSFlexGrid1.Row))
                    Label0(5).Caption = Trim$(Str$(iWert))
                    gbAender = True
                    gbUpdate = True
                    SaveIndirektWK15a
                    giErsetzen = 0
                End If
            Case Is = 13    'Hoch
                If Label0(0).Caption <> "-999" Then
                    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
                    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
                    iWert = Val(Label0(0).Caption)
                    If iWert > 1 Then
                        iWert = iWert - 1
                    End If
                    Label0(4).Caption = Trim$(Str$(iWert))
                    Label0(5).Caption = Trim$(Str$(MSFlexGrid1.Col))
                    gbAender = True
                    gbUpdate = True
                    SaveIndirektWK15a
                    giErsetzen = 0
                End If
            Case Is = 14    'Tief
                If Label0(0).Caption <> "-999" Then
                    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
                    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
                    iWert = Val(Label0(0).Caption)
                    If iWert < MSFlexGrid1.Rows - 1 Then
                        iWert = iWert + 1
                    End If
                    Label0(4).Caption = Trim$(Str$(iWert))
                    Label0(5).Caption = Trim$(Str$(MSFlexGrid1.Col))
                    gbAender = True
                    gbUpdate = True
                    SaveIndirektWK15a
                    giErsetzen = 0
                End If
            Case Is = 15    'Rechts
                If Label0(0).Caption <> "-999" Then
                    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
                    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
                    iWert = Val(Label0(1).Caption)
                    If iWert < MSFlexGrid1.Cols - 1 Then
                        iWert = iWert + 1
                    End If
                    Label0(4).Caption = Trim$(Str$(MSFlexGrid1.Row))
                    Label0(5).Caption = Trim$(Str$(iWert))
                    MSFlexGrid1.Row = Val(Label0(0).Caption)
                    MSFlexGrid1.Col = Val(Label0(1).Caption)
                    gbAender = True
                    gbUpdate = True
                    SaveIndirektWK15a
                    giErsetzen = 0
                End If
        End Select
    Else
        Select Case Index
            
            Case Is = 12    'Links
                iWert = Val(Label0(1).Caption)
                If iWert > 3 Then
                    iWert = iWert - 1
                End If
                Label0(1).Caption = Trim$(Str$(iWert))
                
            Case Is = 13    'Hoch
                iWert = Val(Label0(0).Caption)
                If iWert > 1 Then
                    iWert = iWert - 1
                End If
                Label0(0).Caption = Trim$(Str$(iWert))
            
            Case Is = 14    'Tief
                iWert = Val(Label0(0).Caption)
                If iWert < MSFlexGrid1.Rows - 1 Then
                    iWert = iWert + 1
                End If
                Label0(0).Caption = Trim$(Str$(iWert))
    
            Case Is = 15    'Rechts
                iWert = Val(Label0(1).Caption)
                If iWert < MSFlexGrid1.Cols - 1 Then
                    iWert = iWert + 1
                End If
                Label0(1).Caption = Trim$(Str$(iWert))
                MSFlexGrid1.Row = Val(Label0(0).Caption)
                MSFlexGrid1.Col = Val(Label0(1).Caption)
            
        End Select
    
    End If
    
    If Label0(0).Caption <> "-999" Then
        MSFlexGrid1.Row = Val(Label0(0).Caption)
        MSFlexGrid1.Col = Val(Label0(1).Caption)
        MSFlexGrid1.SetFocus
    Else
        Text1.SetFocus
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten." & Index & " " & iStufe
    
    Fehlermeldung1
    
End Sub
Private Sub SucheInGrid(sArtnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
   
    MSFlexGrid1.Redraw = False
    
    MSFlexGrid1.Row = 0
     
    For j = 0 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Col = 0
        MSFlexGrid1.Row = j
         
        If MSFlexGrid1.Text = sArtnr Then
                    
            MSFlexGrid1.TopRow = j
            MSFlexGrid1.Col = 5
            MSFlexGrid1.Row = j
            MSFlexGrid1.SetFocus
            Exit For
        Else
            MSFlexGrid1.TopRow = 1
            MSFlexGrid1.Row = 1
        End If
    Next j
    MSFlexGrid1.Redraw = True
     
        
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheInGrid"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei      As String
    Dim cSQL        As String
    Dim cPfad       As String
    Dim cDatum      As String
    Dim cLieferant  As String
    Dim iRet        As Integer
    Dim iZufall     As Integer
    Dim cTabelle    As String
    Dim dbBestell   As Database
    Dim cDatname    As String
    Dim sSQL        As String
    Dim lDatum      As Long
    Dim lZaehler    As Long
    Dim lcount      As Long
    Dim lxMal       As Long
    Dim cArtNr      As String
    Dim ctemp       As String
    Dim sCheckLinr  As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Select Case Index
        Case Is = 0
            voreinstellungspeichern
            
            If List2.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
            Else
            
                cdatei = List2.list(List2.ListIndex)
                cLieferant = Mid(cdatei, 27, Len(cdatei) - 26)
                Label2(2).Caption = cLieferant
                cdatei = Trim(Left(cdatei, 13))
                gcWEdatei = ""
                gcWEdatei = cdatei
                Label3(10).Caption = Trim(Right(List2.list(List2.ListIndex), 8))
                
                'check doch mal ob es Budni ist
                'wenn ja dann check mal ob ein Lieferavis vorliegt
                check_Budni_Lieferavis cdatei
                
                
                
                
                
                
                Randomize
                iZufall = 0
                iZufall = Int((999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 999 generieren.
                sZufall = Str(iZufall)
                sZufall = "A" & Trim(sZufall)
                
                cTabelle = cdatei
                
                If NewTableSuchenDBKombi(cTabelle, gdBase) Then
                    If SpalteInTabellegefundenNEW(cTabelle, "BESTELLT", gdBase) Then
                        cSort = ermLastSortierung(cdatei)
                        HoleZwischengespeichertebestelldatei cdatei
                        cArtNr = ermLastfocus(cdatei)
                        
                        SucheInGrid cArtNr
                    Else
                        HoleBestellDateiWK15a cdatei
                    End If
                    
                    'artnr + geliefert speichern
                    
                    
                    loesch "TempZufall" & sZufall
    
                    cSQL = "Create Table TempZufall" & sZufall
                    cSQL = cSQL & " ( ARTNR Long"
                    cSQL = cSQL & ", GELIEFERT Long "
                    cSQL = cSQL & " ) "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Insert into TempZufall" & sZufall
                    cSQL = cSQL & " Select ARTNR,GELIEFERT"
                    cSQL = cSQL & " from " & sZufall & " "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    'Ende artnr + geliefert speichern
                    
                    
                Else
                    cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Delete from LASTFOCUS where TABNAME like '" & cdatei & "*' "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    loeschNEW cdatei, gdBase
                    List2.RemoveItem List2.ListIndex
                    Screen.MousePointer = 0
                    Exit Sub
                
                End If
                
                Text7.Visible = True
                Text5.Visible = True
                Label3(5).Visible = True
                Label3(1).Visible = True
                Command6.Visible = True
                Check2.Visible = True
                Check6.Visible = True
                
                Check4.Visible = True
                
                If Check4.Value = vbChecked Then
                    cboStrichEndlos.Visible = True
                Else
                    cboStrichEndlos.Visible = False
                End If
                
                
                
                Check1.Visible = True
                Text10.Visible = True
                Label3(4).Visible = True
                
                Text5.SetFocus
                
                
                Dim KI As Integer
                Dim cArtXX As String
                
                MSFlexGrid1.Refresh
                
                With MSFlexGrid1
                    For KI = 1 To .Rows - 1
                        .Row = KI
            
                        cArtXX = "0"
                        cArtXX = Val(.TextMatrix(KI, 0))
    
                    Next KI
                End With
        
                
            End If
            
            
            If Val(Label3(6).Caption) > 0 Then
                Text1.Text = ermDepotRabatt1_Lief(CLng(Label3(6).Caption))
                
                If IsNumeric(Text1.Text) = True Then
                    If Text1.Text <> "0" Then
                        Command3_Click
                    End If
                End If
                
            End If
            
        Case Is = 1
            If List2.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
            Else
                lZaehler = 0
                
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                        lZaehler = lZaehler + 1
                    End If
                Next lcount
            
                If lZaehler > 1 Then
                    iRet = MsgBox("Wollen Sie die " & lZaehler & " Bestell-Dateien wirklich löschen?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbYes Then
                    
                        For lcount = 0 To List2.ListCount - 1
                            If List2.Selected(lcount) = True Then
                                cdatei = UCase$(Trim$(Left(List2.list(lcount), 13)))
                                
                                cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
                                gdBase.Execute cSQL, dbFailOnError
                                
                                cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
                                gdBase.Execute cSQL, dbFailOnError
                                
                                loeschNEW cdatei, gdBase
                            End If
                        Next lcount
                    End If
                    
                    LeseInhaltWK15a
                
                Else
                
                    cdatei = UCase$(Trim$(Left(List2.list(List2.ListIndex), 13)))
                
                    iRet = MsgBox("Wollen Sie die Bestell-Datei " & vbCrLf & vbCrLf & cdatei & vbCrLf & vbCrLf & " wirklich löschen?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbYes Then
                        
                        cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
                        gdBase.Execute cSQL, dbFailOnError
                        
                        cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
                        gdBase.Execute cSQL, dbFailOnError
                        
                        loeschNEW cdatei, gdBase
                        
                        List2.RemoveItem List2.ListIndex
                    End If
                
                End If
            End If
        
        Case Is = 2
            Unload frmWK15a
        Case Is = 3 'Kundenbestellungen anzeigen
            KB "GELIEFERT", "INFORMIEREN"
            UpdateKuBestKUNDENSTATUS "INFORMIEREN", "GELIEFERT"
            Command1(3).Visible = False
        Case 4 'Exportieren
        
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
                Exit Sub
            Else
                cDatname = List2.list(List2.ListIndex)
                cDatname = Left(cDatname, 13)
                cDatname = UCase$(Trim$(cDatname))
            End If
            
            cPfad = gcDBPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            cPfad = cPfad & "Bestell\"
            
            Screen.MousePointer = 0
            With cdlopen
                .CancelError = True
                On Error GoTo err
                .DialogTitle = "Datei speichern"
                .Filter = "Access - Dateien (*.mdb)|*.mdb|CSV - Dateien (*.csv)|*.csv"
                .FileName = cPfad & cDatname
               
                .ShowSave
            End With
    
            cPfad = cdlopen.FileName
            
            If FileExists(cPfad) Then
                iRet = MsgBox("Eine gleichnamige Datei ist schon vorhanden, möchten Sie diese überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbNo Then
                    
                    Exit Sub
                Else
                    Kill cPfad
                End If
            End If
            
            If cdlopen.FilterIndex = 1 Then 'access
            
                Set dbBestell = CreateDatabase(cPfad, dbLangGeneral, dbVersion40)
                dbBestell.Close
                
                
                sSQL = "select * into " & cDatname & " in '" & cPfad & "' from " & cDatname
                gdBase.Execute sSQL, dbFailOnError
            
            ElseIf cdlopen.FilterIndex = 2 Then 'csv
            
                Export_inCSV cPfad, cDatname
            
            End If
        
        Case 5 'Importieren
        
            ctemp = "Möchten Sie eine Excel-Tabelle importieren?"
            iRet = MsgBox(ctemp, vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
                excel_import
            Else
                access_import
            End If
            
        Case 6 'Drucken
        
            loeschNEW "PRINTQ", gdBase
            CreateTable "PRINTQ", gdBase
        
            If List2.ListCount > 0 Then
                For lcount = 0 To List2.ListCount - 1
                    cdatei = List2.list(lcount)
                    cSQL = "Insert into PrintQ (Zeile) values ('" & cdatei & "')"
                    gdBase.Execute cSQL, dbFailOnError
                Next lcount
                reportbildschirm "WKL029", "aWKL15ab"
            End If
            
        Case 7 'Zusammenfassen
        
            Dim cLinrPruef      As String
            Dim cNextLinrName   As String
            cLinrPruef = ""
            If List2.ListCount = 0 Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
            Else
                lZaehler = 0
                For lcount = 0 To List2.ListCount - 1
                    If List2.Selected(lcount) = True Then
                        lZaehler = lZaehler + 1
                    End If
                Next lcount
            
                If lZaehler > 1 Then
                    iRet = MsgBox("Wollen Sie die " & lZaehler & " Bestell-Dateien wirklich zusammenlegen?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbYes Then
                    
                        For lcount = 0 To List2.ListCount - 1
                            If List2.Selected(lcount) = True Then
                                cdatei = UCase$(Trim$(Left(List2.list(lcount), 13)))
                                
                                If cLinrPruef <> "" Then
                                    If Mid(cdatei, 2, Len(cdatei) - 2) <> cLinrPruef Then
                                        MsgBox "Bitte gleiche Lieferanten auswählen!", vbInformation, "Winkiss Hinweis:"
                                        Exit For
                                        Exit Sub
                                    End If
                                Else
                                    cLinrPruef = Mid(cdatei, 2, Len(cdatei) - 2)
                                End If
                    
                            End If
                        Next lcount
                        
                        cNextLinrName = NaechsterBestellname(cLinrPruef)
                    End If
                    
                    If cNextLinrName <> "" Then
                    
                        loeschNEW cNextLinrName, gdBase
    
                        sSQL = " Create Table " & cNextLinrName
                        sSQL = sSQL & " ( "
                        sSQL = sSQL & " ARTNR double"
                        sSQL = sSQL & ", LIBESNR Text(13)"
                        sSQL = sSQL & ", BEZEICH Text(35)"
                        sSQL = sSQL & ", LEKPR Double"
                        sSQL = sSQL & ", BESTELLT Long"
                        sSQL = sSQL & ", GELIEFERT Long"
                        sSQL = sSQL & ", BERECHNET Long"
                        sSQL = sSQL & ", LIEF Double"
                        sSQL = sSQL & ", ZEILENRAB Double"
                        sSQL = sSQL & ", ZEILE Double"
                        sSQL = sSQL & ", RECHNRAB Double"
                        sSQL = sSQL & ", RECHN Double"
                        sSQL = sSQL & ", STCK Double"
                        sSQL = sSQL & ", KVKPR1 Double"
                        sSQL = sSQL & ", MOPREIS LONG"
                        sSQL = sSQL & ", LINR LONG"
                        sSQL = sSQL & ", LPZ LONG"
                        sSQL = sSQL & ", BESTVOR Long"
                        sSQL = sSQL & " ) "
                        gdBase.Execute sSQL, dbFailOnError
                    
                        lxMal = 0
                        lDatum = Fix(Now)
                        cDatum = Trim$(Str$(lDatum))
                        For lcount = 0 To List2.ListCount - 1
                            If List2.Selected(lcount) = True Then
                                cdatei = UCase$(Trim$(Left(List2.list(lcount), 13)))
                                
                                If lxMal = 0 Then
                                    'beim ersten mal
                                    If NewTableSuchenDBKombi(cdatei, gdBase) Then
                                        If SpalteInTabellegefundenNEW(cdatei, "BESTELLT", gdBase) Then
                                        
                                            'zwischengespeicherte
                                            sSQL = "Insert into " & cNextLinrName & " "
                                            sSQL = sSQL & " Select * from " & cdatei
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                            sSQL = "Insert into BESTREST "
                                            sSQL = sSQL & "Select LINR, "
                                            sSQL = sSQL & "ARTNR, LEKPR, BESTELLT as BESTVOR, '" & cNextLinrName & ".DBF' as DATEINAME, "
                                            sSQL = sSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
                                            sSQL = sSQL & " from " & cdatei & " where BESTELLT <> 0 "
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                        Else
                                        
                                            sSQL = "Insert into " & cNextLinrName & " "
                                            sSQL = sSQL & " Select ARTNR, BEZEICH, LEKPR, BESTVOR as BESTELLT"
                                            sSQL = sSQL & ", BESTVOR as GELIEFERT, BESTVOR as BERECHNET"
                                            sSQL = sSQL & ", LINR, LIBESNR, KVKPR1,LPZ ,mopreis "
                                            sSQL = sSQL & " from " & cdatei & " where BESTVOR <> 0 "
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                            sSQL = "Insert into BESTREST "
                                            sSQL = sSQL & "Select LINR, "
                                            sSQL = sSQL & "ARTNR, LEKPR, BESTVOR, '" & cNextLinrName & ".DBF' as DATEINAME, "
                                            sSQL = sSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
                                            sSQL = sSQL & " from " & cdatei & " where BESTVOR <> 0 "
                                            gdBase.Execute sSQL, dbFailOnError
                                        End If
                                    End If
                                Else
                                    'ab dem 2. Mal
                                    
                                    'bestelldatei
                                    If NewTableSuchenDBKombi(cdatei, gdBase) Then
                                        If SpalteInTabellegefundenNEW(cdatei, "BESTELLT", gdBase) Then
                                        
                                            'erst update dann insert
                                            sSQL = "update " & cNextLinrName & " inner join " & cdatei & " on " & cNextLinrName & ".artnr = " & cdatei & ".artnr "
                                            sSQL = sSQL & " set " & cNextLinrName & ".BESTELLT = " & cNextLinrName & ".BESTELLT + " & cdatei & ".BESTELLT "
                                            sSQL = sSQL & " , " & cNextLinrName & ".GELIEFERT = " & cNextLinrName & ".GELIEFERT + " & cdatei & ".BESTELLT "
                                            sSQL = sSQL & " , " & cNextLinrName & ".BERECHNET = " & cNextLinrName & ".BERECHNET + " & cdatei & ".BESTELLT "
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                            sSQL = "Insert into " & cNextLinrName & " Select * "
                                            sSQL = sSQL & " from " & cdatei
                                            sSQL = sSQL & " where not artnr in (select artnr from " & cNextLinrName & " )"
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                            'bestrest
                                            'erst update dann insert
                                            sSQL = "Update BESTREST inner join " & cdatei & " on BESTREST.artnr = " & cdatei & ".artnr "
                                            sSQL = sSQL & " set BESTREST.bestvor = BESTREST.bestvor + " & cdatei & ".BESTELLT "
                                            sSQL = sSQL & " where Bestrest.BEST_DATUM = " & cDatum
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                            sSQL = "Insert into BESTREST "
                                            sSQL = sSQL & "Select LINR, "
                                            sSQL = sSQL & "ARTNR, LEKPR, BESTELLT as BESTVOR, '" & cNextLinrName & ".DBF' as DATEINAME, "
                                            sSQL = sSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
                                            sSQL = sSQL & " from " & cdatei & " where BESTELLT <> 0 "
                                            sSQL = sSQL & " and not artnr in (select artnr from BESTREST where BEST_DATUM = " & cDatum & ")"
                                            gdBase.Execute sSQL, dbFailOnError
                                    
                                        Else
                                            'erst update dann insert
                                            sSQL = "update " & cNextLinrName & " inner join " & cdatei & " on " & cNextLinrName & ".artnr = " & cdatei & ".artnr "
                                            sSQL = sSQL & " set " & cNextLinrName & ".BESTELLT = " & cNextLinrName & ".BESTELLT + " & cdatei & ".bestvor "
                                            sSQL = sSQL & " , " & cNextLinrName & ".GELIEFERT = " & cNextLinrName & ".GELIEFERT + " & cdatei & ".bestvor "
                                            sSQL = sSQL & " , " & cNextLinrName & ".BERECHNET = " & cNextLinrName & ".BERECHNET + " & cdatei & ".bestvor "
                                            gdBase.Execute sSQL, dbFailOnError
                                        
                                            sSQL = "Insert into " & cNextLinrName & " "
                                            sSQL = sSQL & " Select ARTNR, BEZEICH, LEKPR, BESTVOR as BESTELLT"
                                            sSQL = sSQL & ", BESTVOR as GELIEFERT, BESTVOR as BERECHNET"
                                            sSQL = sSQL & ", LINR, LIBESNR, KVKPR1,LPZ ,mopreis "
                                            sSQL = sSQL & " from " & cdatei
                                            sSQL = sSQL & " where not artnr in (select artnr from " & cNextLinrName & " )"
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                            'bestrest
                                            'erst update dann insert
                                            sSQL = "Update BESTREST inner join " & cdatei & " on BESTREST.artnr = " & cdatei & ".artnr "
                                            sSQL = sSQL & " set BESTREST.bestvor = BESTREST.bestvor + " & cdatei & ".bestvor "
                                            sSQL = sSQL & " where Bestrest.BEST_DATUM = " & cDatum
                                            gdBase.Execute sSQL, dbFailOnError
                                            
                                            sSQL = "Insert into BESTREST "
                                            sSQL = sSQL & "Select LINR, "
                                            sSQL = sSQL & "ARTNR, LEKPR, BESTVOR, '" & cNextLinrName & ".DBF' as DATEINAME, "
                                            sSQL = sSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
                                            sSQL = sSQL & " from " & cdatei & " where BESTVOR <> 0 "
                                            sSQL = sSQL & " and not artnr in (select artnr from BESTREST where BEST_DATUM = " & cDatum & ")"
                                            gdBase.Execute sSQL, dbFailOnError
                                        End If
                                    End If
                                End If
                                
                                cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
                                gdBase.Execute cSQL, dbFailOnError
                                
                                cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
                                gdBase.Execute cSQL, dbFailOnError
                                
                                loeschNEW cdatei, gdBase
                                
                                lxMal = lxMal + 1
                                
                            End If
                        Next lcount
                    
                    End If
                    
                    Screen.MousePointer = 0
                    neuFildatschreiben
                    Screen.MousePointer = 11
                    
                    LeseInhaltWK15a
                
                Else
                    MsgBox "Bitte mindestens 2 Dateien auswählen!", vbInformation, "Winkiss Hinweis:"
                End If
            End If
            
        Case 8 'als Bestellvorschlag
        
            anzeige "normal", "", Label9
        
            Dim cDatnameAlt As String
        
            If List2.ListIndex < 0 Then
                MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
                Exit Sub
            Else
                cDatname = List2.list(List2.ListIndex)
                cDatname = Left(cDatname, 13)
                cDatname = UCase$(Trim$(cDatname))
                cDatnameAlt = cDatname
            End If
            
            cPfad = gcPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
            
            cPfad = cPfad & "Kissapp.mdb"
            
            cDatname = "X" & Right(cDatname, Len(cDatname) - 1)
            
            If NewTableSuchenDBKombi(cDatname, gdApp) = True Then
                iRet = MsgBox("Eine gleichnamige Datei ist schon vorhanden, möchten Sie diese überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbNo Then
                    anzeige "rot", "Abbruch", Label9
                    Screen.MousePointer = 0
                    Exit Sub
                Else
                    loeschNEW cDatname, gdApp
                End If
            End If
            
            sSQL = "select * into " & cDatname & " in '" & cPfad & "' from " & cDatnameAlt
            gdBase.Execute sSQL, dbFailOnError
            
            anzeige "normal", "erfolgreich als Bestellvorschlag bereitgestellt", Label9
            
    End Select
    
    Screen.MousePointer = 0
    
err:
Exit Sub

LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub
Private Sub access_import()
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim dbBestell   As Database
    Dim lAnzTable   As Long
    Dim cDatname    As String
    Dim cDatum      As String
    Dim iRet        As Integer
    Dim sSQL        As String
    Dim lDatum      As Long
    Dim lcount      As Long

    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "Bestell\"
    
    Screen.MousePointer = 0
    With cdlopen
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Datei importieren"
        .Filter = "Access - Dateien (*.mdb)|*.mdb"
        .FileName = cPfad & cDatname & ".mdb"
        .ShowOpen
    End With

    cPfad = cdlopen.FileName
    
    Set dbBestell = OpenDatabase(cPfad, False, False)
    dbBestell.TableDefs.Refresh

    lAnzTable = dbBestell.TableDefs.Count

    For lcount = 0 To lAnzTable - 1
        cDatname = dbBestell.TableDefs(lcount).name
        If Left(UCase(cDatname), 1) = "Q" Then
            If NewTableSuchenDBKombi(cDatname, gdBase) Then
                iRet = MsgBox("Eine gleichnamige Datei ist schon vorhanden, möchten Sie diese überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbNo Then
                    
                    Exit Sub
                Else
                    loeschNEW cDatname, gdBase
                End If
            Else
             
            End If
             
            cPfad = gcDBPfad
            If Right(cPfad, 1) <> "\" Then
                cPfad = cPfad & "\"
            End If
             
            TransferTab dbBestell, cPfad & "Kissdata.mdb", cDatname
            
            'BESTREST füllen
            sSQL = "Delete from BESTREST where DATEINAME = '" & cDatname & ".DBF'"
            gdBase.Execute sSQL, dbFailOnError
            
            lDatum = Fix(Now)
            cDatum = Trim$(Str$(lDatum))
    
            sSQL = "Insert into BESTREST "
            sSQL = sSQL & "Select LINR, "
            sSQL = sSQL & "ARTNR, LEKPR, BESTVOR, '" & cDatname & ".DBF' as DATEINAME, "
            sSQL = sSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
            sSQL = sSQL & " from " & cDatname & " where BESTVOR <> 0 "
            gdBase.Execute sSQL, dbFailOnError
            
            dbBestell.Close
            Screen.MousePointer = 0
            neuFildatschreiben
            Screen.MousePointer = 11
            Exit For
        End If
    Next lcount
    LeseInhaltWK15a
    
err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "access_import"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub excel_import()
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim dbExcel     As Database
    Dim lAnzTable   As Long
    Dim cDatname    As String
    Dim cDatum      As String
    Dim iRet        As Integer
    Dim sSQL        As String
    Dim lDatum      As Long
    Dim lcount      As Long
    Dim gsExcel50   As String
    Dim rsrs        As Recordset
    Dim rsKU        As Recordset
    Dim lLinr       As Long
    
    loeschNEW "BESTIMP", gdBase
    CreateTableT2 "BESTIMP", gdBase
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "IN\"
    
    Screen.MousePointer = 0
    With cdlopen
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Datei importieren"
        .Filter = "Excel - Dateien (*.xls)|*.xls"
        .FileName = cPfad & cDatname & ".xls"
        .ShowOpen
    End With

    cPfad = cdlopen.FileName
    gsExcel50 = "Excel 5.0;"
        
    Set dbExcel = OpenDatabase(cPfad, 0, 0, gsExcel50)
    Set rsKU = gdBase.OpenRecordset("BESTIMP")
    Set rsrs = dbExcel.OpenRecordset("Import$")
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsKU.AddNew
            
            rsKU!artnr = rsrs!KISSARTNR
            rsKU!Menge = rsrs!Menge
            rsKU!ekpr = rsrs!ekpr
            rsKU.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    rsKU.Close: Set rsKU = Nothing
    
    
    lLinr = checkLinrForKISS(Label9)
    
    If lLinr > 0 Then
        Create_Bestell_Datei lLinr
    Else
        anzeige "rot", "Abbruch, kein Lieferant ausgewählt", Label9
    End If
    
err:
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "excel_import"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Export_inCSV(sPfadUndDatname As String, sDatei As String)
   On Error GoTo LOKAL_ERROR

    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim cSatz       As String
    Dim rsrs        As Recordset
    Dim cSQL        As String
    
    Dim bSpalteNameBESTELLT As Boolean
    
    
    If SpalteInTabellegefundenNEW(sDatei, "BESTELLT", gdBase) Then
        bSpalteNameBESTELLT = True
    Else
        bSpalteNameBESTELLT = False
    End If
    
    Kill sPfadUndDatname
    iFileNr = FreeFile
    Open sPfadUndDatname For Binary As #iFileNr
    
    cSQL = "Select * from  " & sDatei
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!LIBESNR) Then
            
                cSatz = ""
                cSatz = cSatz & rsrs!LIBESNR & ";"
                
                If bSpalteNameBESTELLT = True Then
                    cSatz = cSatz & ";"
                    cSatz = cSatz & rsrs!BESTELLT & vbCrLf
                Else
                    cSatz = cSatz & rsrs!EAN & ";"
                    cSatz = cSatz & rsrs!BESTVOR & vbCrLf
                End If
                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cSatz
                
            End If
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close
    Close iFileNr

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Export_inCSV"
        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
        Fehlermeldung1
    End If
End Sub
Private Function NaechsterBestellname(cLinr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim bgefunden As Boolean
    Dim ctempa As String

    NaechsterBestellname = ""
    
    lcount = 65
    bgefunden = True
    Do While bgefunden
        NaechsterBestellname = "Q" & cLinr & Chr$(lcount)
        If NewTableSuchenDBKombi("Q" & cLinr & Chr$(lcount), gdBase) Then
            bgefunden = True
            lcount = lcount + 1
        Else
            bgefunden = False
        End If
    Loop
        
    If lcount > 89 Then
        
        ctempa = "Die Vergabe eines Dateinamens ist gescheitert." & vbCrLf
        ctempa = ctempa & "Löschen Sie erledigte Bestellungen im Wareneingang aus Bestellung!"
        MsgBox ctempa, vbOKOnly + vbInformation, "Winkiss Hinweis:"
        Screen.MousePointer = 0
        Exit Function
    End If

            
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NaechsterBestellname"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei  As String
    Dim iRet    As Integer
    Dim cPfad   As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Select Case Index
        Case Is = 0
        
            If IsAktionZulaessig("Lieferung übernehmen") = False Then
                Exit Sub
            End If

            iRet = MsgBox("Die angezeigten Daten in den Artikelbestand übernehmen?", vbQuestion + vbYesNo, "Winkiss Frage:")
            If iRet = vbYes Then
                Screen.MousePointer = 11

                If gbDIFFPROT Then
                    DiffProt_drucken
                End If
                
                If gbUEBERPROT Then
                    druckendgültigTeil1
                End If
                
               
                
                
                EingangDerArtikel
                
                If gbUEBERPROT Then
                    druckendgültigTeil2
                End If
                
                If gbGescheitert = True Then
                    gbGescheitert = False
                    Frame2.Visible = True
                    Frame0.Visible = True
                    Frame1.Enabled = False
                Else
                    LeseInhaltWK15a
                    
                    Label2(1).Caption = ""
                    Label2(2).Caption = ""
                    Frame1.Visible = True
                    Frame1.Enabled = True
                    Frame2.Visible = False
                    Frame0.Visible = False
                End If
            End If

            AktionAustragen "Lieferung übernehmen"
        Case Is = 1
            iRet = MsgBox("Möchten Sie wirklich den Wareneingang verlassen?", vbQuestion + vbYesNo, "Winkiss Frage:")
            
            If iRet = vbYes Then
                frame2close
            End If
        Case Is = 2
            Text2.Text = ""
            Text3.Text = ""
            Text4.Text = ""
            Text6.Text = ""
            List4.Clear
            List4.Visible = True
            List3.Clear
            List3.Visible = True
            
            cmdAnfuegen_Click 0
            fraArtAnfuegen.Visible = True
            Text4.SetFocus
            
            
        Case Is = 3 'Zwischenspeichern
        
            MSFlexGrid1_SelChange
            
            Zwischenspeichern_neu
            
            Screen.MousePointer = 11
            neuFildatschreiben
            
            Screen.MousePointer = 0
            
            LeseInhaltWK15a
            
            frame2close
            
            
        Case Is = 4
        
            'zum Etikettendruck
            schreibeEtiketten
            frmWKL30.Show 1
        Case 16
        
            Text8.Text = Format(Datumschreiben11a(3000, 9000), "DD.MM.YY")
            
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command2_Click"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub schreibeEtiketten()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lFil As Long
    Dim cArtNr As String
    Dim lAnzahl   As Long
    
    lFil = CLng(gcFilNr)
    
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Visible = False
    
    For lrow = 2 To MSFlexGrid1.Rows - 1

    
        MSFlexGrid1.Row = lrow
        MSFlexGrid1.Col = 0
        cArtNr = MSFlexGrid1.Text
        
        MSFlexGrid1.Col = 5
        lAnzahl = Val(MSFlexGrid1.Text)
        
        schreibeWKEtidru cArtNr, lAnzahl, lFil
        
        
    Next lrow
    
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Redraw = True

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "schreibeEtiketten"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub frame2close()
On Error GoTo LOKAL_ERROR

    Frame0.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    Frame1.Enabled = True
    
    Text7.Visible = False
    Label3(5).Visible = False
    Text5.Visible = False
    Label3(1).Visible = False
    Command6.Visible = False
    Check2.Visible = False
    Check6.Visible = False
    
    cboStrichEndlos.Visible = False
    Check4.Visible = False
    Check1.Visible = False
    Text10.Visible = False
    Label3(4).Visible = False
    
    lblUeberschrift(1).Caption = ""
    Text1.Text = ""
    MSFlexGrid1.Clear
    
    If sZufall <> "" Then
        loeschNEW sZufall, gdBase
    End If
    
'    LeseInhaltWK15a
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "frame2close"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zwischenspeichern()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsQZW       As Recordset
    Dim i           As Integer
    Dim j           As Integer
    Dim cTabelle    As String
    Dim cdatei      As String
    Dim cArtNr      As String
    Dim cDelLinr      As String
    
    cdatei = List2.list(List2.ListIndex)
    cdatei = Trim(Left(cdatei, 10))
    cdatei = UCase$(cdatei)
    
    cDelLinr = Left(cdatei, Len(cdatei) - 1)
    cDelLinr = Right(cDelLinr, Len(cDelLinr) - 1)
    
    sSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    cTabelle = "Q" & cAnfuLinr & "Z"
    gcWEdatei = "Q" & cAnfuLinr & "Z"
    
    
    sSQL = "Update BESTREST set DATEINAME = '" & cTabelle & "' where DATEINAME like '" & cdatei & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW cTabelle, gdBase
    
    sSQL = " Create Table " & cTabelle
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR double"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LEKPR Double"
    sSQL = sSQL & ", BESTELLT Long"
    sSQL = sSQL & ", GELIEFERT Long"
    sSQL = sSQL & ", BERECHNET Long"
    sSQL = sSQL & ", LIEF Double"
    sSQL = sSQL & ", ZEILENRAB Double"
    sSQL = sSQL & ", ZEILE Double"
    sSQL = sSQL & ", RECHNRAB Double"
    sSQL = sSQL & ", RECHN Double"
    sSQL = sSQL & ", STCK Double"
    sSQL = sSQL & ", KVKPR1 Double"
    sSQL = sSQL & ", MOPREIS LONG"
    sSQL = sSQL & ", LINR LONG"
    sSQL = sSQL & ", LPZ LONG"
    sSQL = sSQL & ", BESTVOR Long"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from TABDATUM where TABNAME like '" & cTabelle & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into TABDATUM (Tabname,Tabdate) values"
    sSQL = sSQL & " ( '" & cTabelle & "','" & DateValue(Now) & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsQZW = gdBase.OpenRecordset(cTabelle, dbOpenTable)
    
    MSFlexGrid1.Redraw = False
    
    
    With MSFlexGrid1
        For i = 1 To .Rows - 1
            .Row = i
            
            cArtNr = "0"
            cArtNr = Val(.TextMatrix(i, 0))
            
            If isArtnrEnthalten(cArtNr, cTabelle) = True Then
            
'''                MsgBox cArtNr für Rühle 8.5.15
            
            Else

            
                If cArtNr <> "0" Then
                
                    rsQZW.AddNew
                    For j = 0 To .Cols - 1
                        .Col = j
                        If j = 0 Then
                            rsQZW(j) = CDbl(.Text)
                        ElseIf j < 3 Then
                            rsQZW(j) = Trim(.Text)
                        Else
                            If .Text = "" Then
                                rsQZW(j) = 0
                            Else
                                If Val(.Text) > 999999 Then
                                    rsQZW(j) = 0
                                Else
                                    If IsNumeric(.Text) Then
                                        rsQZW(j) = CDbl(.Text)
                                    Else
                                        rsQZW(j) = Val(.Text)
                                    End If
                                End If
                            End If
                            
                        End If
                    Next j
                    rsQZW.Update
                End If
            End If
        Next i
    End With
    
    
    
    
    
     
    MSFlexGrid1.Redraw = True
    
    rsQZW.Close
    
    sSQL = "Delete from " & cTabelle & " where artnr not in (Select artnr from artlief where linr = " & cDelLinr & ")"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Update " & cTabelle & " set BESTVOR = Bestellt"
    gdBase.Execute sSQL, dbFailOnError
    
    If cdatei <> cTabelle Then
        loesch cdatei
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zwischenspeichern"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Zwischenspeichern_neu()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsQZW       As Recordset
    Dim i           As Integer
    Dim j           As Integer

    Dim cdatei      As String
    Dim cArtNr      As String
    Dim cDelLinr    As String
    
    cdatei = List2.list(List2.ListIndex)
    cdatei = Trim(Left(cdatei, 10))
    cdatei = UCase$(cdatei)
    
    cDelLinr = Left(cdatei, Len(cdatei) - 1)
    cDelLinr = Right(cDelLinr, Len(cDelLinr) - 1)
    
    sSQL = "Update TABDATUM set Kurzinfo = 'z' where TABNAME like '" & cdatei & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("LASTFOCUS", gdBase) = False Then
        CreateTableT2 "LASTFOCUS", gdBase
    End If
    
    cArtNr = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))
    
    sSQL = "Delete from LastFocus where TABNAME like '" & cdatei & "*' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LastFocus (Artnr,TABNAME,Sortierung) values (" & cArtNr & ",'" & cdatei & "','" & cSort & "') "
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW cdatei, gdBase

    sSQL = " Create Table " & cdatei
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR double"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", LEKPR Double"
    sSQL = sSQL & ", BESTELLT Long"
    sSQL = sSQL & ", GELIEFERT Long"
    sSQL = sSQL & ", BERECHNET Long"
    sSQL = sSQL & ", LIEF Double"
    sSQL = sSQL & ", ZEILENRAB Double"
    sSQL = sSQL & ", ZEILE Double"
    sSQL = sSQL & ", RECHNRAB Double"
    sSQL = sSQL & ", RECHN Double"
    sSQL = sSQL & ", STCK Double"
    sSQL = sSQL & ", KVKPR1 Double"
    sSQL = sSQL & ", MOPREIS LONG"
    sSQL = sSQL & ", LINR LONG"
    sSQL = sSQL & ", LPZ LONG"
    sSQL = sSQL & ", BESTVOR Long"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError

    Set rsQZW = gdBase.OpenRecordset(cdatei, dbOpenTable)
    MSFlexGrid1.Redraw = False

    With MSFlexGrid1
        For i = 1 To .Rows - 1
            .Row = i

            cArtNr = "0"
            cArtNr = Val(.TextMatrix(i, 0))
            
            If isArtnrEnthalten(cArtNr, cdatei) = True Then
            
'''                MsgBox cArtNr für Rühle 8.5.15
            
            Else

                If cArtNr <> "0" Then
    
                    rsQZW.AddNew
                    For j = 0 To .Cols - 1
                        .Col = j
                        If j = 0 Then
                            rsQZW(j) = CDbl(.Text)
                        ElseIf j < 3 Then
                            rsQZW(j) = Trim(.Text)
                        Else
                            If .Text = "" Then
                                rsQZW(j) = 0
                            Else
                                If Val(.Text) > 999999 Then
                                    rsQZW(j) = 0
                                Else
                                    If IsNumeric(.Text) Then
                                        rsQZW(j) = CDbl(.Text)
                                    Else
                                        rsQZW(j) = Val(.Text)
                                    End If
                                End If
                            End If
    
                        End If
                    Next j
                    rsQZW.Update
                End If
                
            End If
        Next i
    End With

    MSFlexGrid1.Redraw = True

    rsQZW.Close
    
    
    sSQL = "Delete from " & cdatei & " where artnr not in (Select artnr from artlief where linr = " & cDelLinr & ")"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & cdatei & " set BESTVOR = Bestellt"
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zwischenspeichern_neu"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim dWert As Double
    Dim cWert As String
    
    Screen.MousePointer = 11
    
    If Trim$(Text1.Text) = "" Then
        MsgBox "Bitte den Zeilenrabatt eingeben!", vbCritical, "STOP"
        Text1.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    Else
        cWert = Text1.Text
        cWert = fnMoveComma2Point$(cWert)
        dWert = Val(cWert)
    End If
    
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Visible = False
    
    For lrow = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = lrow
        Label0(0).Caption = Trim$(Str$(lrow))
        MSFlexGrid1.Col = 0
        Label0(2).Caption = MSFlexGrid1.Text
        MSFlexGrid1.Col = 8
        Label0(1).Caption = 8
        MSFlexGrid1.Text = Format$(dWert, "#####0.00")
        gbAender = True
        MSFlexGrid1_SelChange
    Next lrow
    
    MoveBestell2GridWK15a cSort
    
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Redraw = True
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command4_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim cPfad   As String
    Dim cFeld   As String
    Dim cFeld1  As String
    Dim cFeld2  As String
    Dim cFeld3  As String
    Dim ctmp    As String
    
    Dim dWert   As Double
    
    Dim lRows   As Long
    Dim lCols   As Long
    Dim lrow    As Long
    Dim lcol    As Long
    Dim iRet    As Integer
    Dim cArtNr  As String
    
    Screen.MousePointer = 11
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    loeschNEW "DRU_TEMP", gdBase
    
    cSQL = "Create Table DRU_TEMP "
    cSQL = cSQL & "("
    cSQL = cSQL & "  BESTELLDAT Text(20)"
    cSQL = cSQL & ", BESTELLUNG Text(20)"
    cSQL = cSQL & ", LIEFERANT Text(35)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", LEKPR Double"
    cSQL = cSQL & ", BESTELLT Long"
    cSQL = cSQL & ", GELIEFERT Long"
    cSQL = cSQL & ", BERECHNET Long"
    cSQL = cSQL & ", LIEF Double"
    cSQL = cSQL & ", ZEILENRAB Double"
    cSQL = cSQL & ", ZEILE Double"
    cSQL = cSQL & ", RECHNRAB Double"
    cSQL = cSQL & ", RECHN Double"
    cSQL = cSQL & ", STCK Double"
    cSQL = cSQL & ", KVKPR1 Double"
    cSQL = cSQL & ", EAN Text(13)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ", LAGERP Long"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld1 = Trim$(Mid(cFeld, 15, 10))
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld2 = Trim$(Mid(cFeld, 1, 10))
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld3 = Trim$(Mid(cFeld, 25, 35))
    
    lRows = MSFlexGrid1.Rows
    lCols = MSFlexGrid1.Cols
    MSFlexGrid1.Redraw = False
    
    For lrow = 1 To lRows - 1
        MSFlexGrid1.Row = lrow
        
        cSQL = "Insert into DRU_TEMP "
        cSQL = cSQL & "( BESTELLDAT"
        cSQL = cSQL & ", BESTELLUNG"
        cSQL = cSQL & ", LIEFERANT"
        cSQL = cSQL & ", ARTNR"
        cSQL = cSQL & ", LIBESNR"
        cSQL = cSQL & ", BEZEICH"
        cSQL = cSQL & ", LEKPR"
        cSQL = cSQL & ", BESTELLT"
        cSQL = cSQL & ", GELIEFERT"
        cSQL = cSQL & ", BERECHNET"
        cSQL = cSQL & ", LIEF"
        cSQL = cSQL & ", ZEILENRAB"
        cSQL = cSQL & ", ZEILE"
        cSQL = cSQL & ", RECHNRAB"
        cSQL = cSQL & ", RECHN"
        cSQL = cSQL & ", STCK"
        cSQL = cSQL & ", KVKPR1"
        cSQL = cSQL & ")"
        cSQL = cSQL & " values ("
        
        cSQL = cSQL & "'" & cFeld1 & "', "
        cSQL = cSQL & "'" & cFeld2 & "', "
        cSQL = cSQL & "'" & cFeld3 & "', "
        
        cArtNr = MSFlexGrid1.TextMatrix(lrow, 0)      'ARTNR
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 0)      'ARTNR
        cSQL = cSQL & "" & ctmp & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 1)      'LIBESNR
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 2)      'BEZEICH
        ctmp = SwapStr(ctmp, "'", "''")
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 3)      'LEKPR
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 4)      'BESTELLT
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 5)      'GELIEFERT
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 6)      'BERECHNET
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 7)      'LIEFSUMME
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 8)      'ZEILENRABATT
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 9)      'ZEILENSUMME
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 10)     'RECHNUNGSRABATT
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 11)     'RECHNUNGSSUMME
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 12)     'STÜCKPREIS
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 13)     'KASSENVK
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ") "
        
        If cArtNr = "" Then
            ctmp = ""
            cFeld = ""
            cSQL = ""
        Else
            
            'MsgBox cSQL
            gdBase.Execute "Delete from DRU_TEMP where artnr = " & cArtNr & " ", dbFailOnError
            cArtNr = ""
            gdBase.Execute cSQL, dbFailOnError
            ctmp = ""
            cFeld = ""
            cSQL = ""
        End If
        
    Next lrow
    MSFlexGrid1.Redraw = True
    
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.ean = artikel.ean "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.ean = artikel.ean2 "
    cSQL = cSQL & " where Dru_temp.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.ean = artikel.ean3 "
    cSQL = cSQL & " where Dru_temp.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.vkpr = artikel.vkpr "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update DRU_TEMP inner join LAGERPLATZ on DRU_TEMP.ARTNR = LAGERPLATZ.ARTNR "
    cSQL = cSQL & " set DRU_TEMP.LAGERP = LAGERPLATZ.LAGERP "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update DRU_TEMP set LAGERP = 0 where lagerp is null "
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    iRet = MsgBox("Möchten Sie die Warenbestellung im Hochformat drucken?", vbYesNo + vbDefaultButton1 + vbQuestion, "Winkiss Frage:")
    If iRet = vbYes Then
        Screen.MousePointer = 11
        reportbildschirm "WKL029", "aWKL15ad"
    Else
        Screen.MousePointer = 11
        reportbildschirm "WKL029", "aWKL15aa"
    End If
    
    Screen.MousePointer = 0
            
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command4_Click"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub DiffProt_drucken()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim cPfad   As String
    Dim cFeld   As String
    Dim cFeld1  As String
    Dim cFeld2  As String
    Dim cFeld3  As String
    Dim ctmp    As String
    
    Dim dWert   As Double
    
    Dim lRows   As Long
    Dim lCols   As Long
    Dim lrow    As Long
    Dim lcol    As Long
    Dim cArtNr  As String
    
    Screen.MousePointer = 11
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    loeschNEW "DRU_TEMP", gdBase
    
    cSQL = "Create Table DRU_TEMP "
    cSQL = cSQL & "("
    cSQL = cSQL & "  BESTELLDAT Text(20)"
    cSQL = cSQL & ", BESTELLUNG Text(20)"
    cSQL = cSQL & ", LIEFERANT Text(35)"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", LIBESNR Text(13)"
    cSQL = cSQL & ", BEZEICH Text(35)"
    cSQL = cSQL & ", LEKPR Double"
    cSQL = cSQL & ", BESTELLT Long"
    cSQL = cSQL & ", GELIEFERT Long"
    cSQL = cSQL & ", BERECHNET Long"
    cSQL = cSQL & ", LIEF Double"
    cSQL = cSQL & ", ZEILENRAB Double"
    cSQL = cSQL & ", ZEILE Double"
    cSQL = cSQL & ", RECHNRAB Double"
    cSQL = cSQL & ", RECHN Double"
    cSQL = cSQL & ", STCK Double"
    cSQL = cSQL & ", KVKPR1 Double"
    cSQL = cSQL & ", EAN Text(13)"
    cSQL = cSQL & ", VKPR Double"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld1 = Trim$(Mid(cFeld, 15, 10))
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld2 = Trim$(Mid(cFeld, 1, 10))
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld3 = Trim$(Mid(cFeld, 25, 35))
            
    lRows = MSFlexGrid1.Rows
    lCols = MSFlexGrid1.Cols
    MSFlexGrid1.Redraw = False
    
    For lrow = 1 To lRows - 1
        MSFlexGrid1.Row = lrow
        
        cSQL = "Insert into DRU_TEMP "
        cSQL = cSQL & "( BESTELLDAT"
        cSQL = cSQL & ", BESTELLUNG"
        cSQL = cSQL & ", LIEFERANT"
        cSQL = cSQL & ", ARTNR"
        cSQL = cSQL & ", LIBESNR"
        cSQL = cSQL & ", BEZEICH"
        cSQL = cSQL & ", LEKPR"
        cSQL = cSQL & ", BESTELLT"
        cSQL = cSQL & ", GELIEFERT"
        cSQL = cSQL & ", BERECHNET"
        cSQL = cSQL & ", LIEF"
        cSQL = cSQL & ", ZEILENRAB"
        cSQL = cSQL & ", ZEILE"
        cSQL = cSQL & ", RECHNRAB"
        cSQL = cSQL & ", RECHN"
        cSQL = cSQL & ", STCK"
        cSQL = cSQL & ", KVKPR1"
        cSQL = cSQL & ")"
        cSQL = cSQL & " values ("
        
        cSQL = cSQL & "'" & cFeld1 & "', "
        cSQL = cSQL & "'" & cFeld2 & "', "
        cSQL = cSQL & "'" & cFeld3 & "', "
        
        cArtNr = MSFlexGrid1.TextMatrix(lrow, 0)      'ARTNR
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 0)      'ARTNR
        cSQL = cSQL & "" & ctmp & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 1)      'LIBESNR
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 2)      'BEZEICH
        ctmp = SwapStr(ctmp, "'", "''")
        cSQL = cSQL & "'" & ctmp & "', "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 3)      'LEKPR
        cFeld = fnMoveComma2Point$(ctmp)
         If cFeld = "" Then cFeld = "0"
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 4)      'BESTELLT
        cFeld = fnMoveComma2Point$(ctmp)
         If cFeld = "" Then cFeld = "0"
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 5)      'GELIEFERT
        cFeld = fnMoveComma2Point$(ctmp)
        
        If cFeld = "" Then cFeld = "0"
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 6)      'BERECHNET
        cFeld = fnMoveComma2Point$(ctmp)
        If cFeld = "" Then cFeld = "0"
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 7)      'LIEFSUMME
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 8)      'ZEILENRABATT
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 9)      'ZEILENSUMME
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 10)     'RECHNUNGSRABATT
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 11)     'RECHNUNGSSUMME
        cFeld = fnMoveComma2Point$(ctmp)
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 12)     'STÜCKPREIS
        cFeld = fnMoveComma2Point$(ctmp)
        If cFeld = "" Then cFeld = "0"
        cSQL = cSQL & cFeld & ", "
        
        ctmp = MSFlexGrid1.TextMatrix(lrow, 13)     'KASSENVK
        cFeld = fnMoveComma2Point$(ctmp)
        If cFeld = "" Then cFeld = "0"
        
        cSQL = cSQL & cFeld & ") "
        
        
        If cArtNr = "" Then
            ctmp = ""
            cFeld = ""
            cSQL = ""
        Else
            
            'MsgBox cSQL
            gdBase.Execute "Delete from DRU_TEMP where artnr = " & cArtNr & " ", dbFailOnError
            cArtNr = ""
            gdBase.Execute cSQL, dbFailOnError
            ctmp = ""
            cFeld = ""
            cSQL = ""
        End If
        
        
        
        
        
        
    Next lrow
    
    MSFlexGrid1.Redraw = True
    Screen.MousePointer = 0
    
    cSQL = "Delete from DRU_TEMP where bestellt = geliefert "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.ean = artikel.ean "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.ean = artikel.ean2 "
    cSQL = cSQL & " where Dru_temp.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.ean = artikel.ean3 "
    cSQL = cSQL & " where Dru_temp.ean is null"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update DRU_TEMP inner join artikel on dru_temp.artnr = artikel.artnr"
    cSQL = cSQL & " set Dru_temp.vkpr = artikel.vkpr "
    gdBase.Execute cSQL, dbFailOnError
    
    reportbildschirm "WKL029", "aWKL15ac"
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DiffProt_drucken"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Command5_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim iRet As Integer
    
    
    iRet = MsgBox("Wirklich?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
    If iRet = vbYes Then
        Select Case Index
            Case Is = 0
                cSQL = "Update " & sZufall & " Set GELIEFERT = 0, BERECHNET = 0 "
            Case Is = 1
                cSQL = "Update " & sZufall & " Set GELIEFERT = BESTELLT, BERECHNET = BESTELLT "
        End Select
        gdBase.Execute cSQL, dbFailOnError
        
        
        MSFlexGrid1.Redraw = False
        
        MoveBestell2GridWK15a cSort
        
        If gbAutoZwsp Then
            MSFlexGrid1.Redraw = False
            MSFlexGrid1_SelChange
            MSFlexGrid1.Redraw = False
            Zwischenspeichern
            
            
            Screen.MousePointer = 11
            MSFlexGrid1.Redraw = False
            neuFildatschreiben
            Screen.MousePointer = 0
        End If
        
        MSFlexGrid1.TopRow = 1
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 5
        MSFlexGrid1.SetFocus
        
        MSFlexGrid1.Redraw = True
        
    End If
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim llaeng      As Long
    Dim ctmp        As String
    Dim sArtnr      As String
    Dim rsArt       As Recordset
    Dim sBez        As String
    Dim i           As Integer
    Dim lRows       As Long
    Dim bFound      As Boolean
    Dim acArtNr(1) As String
    Dim acAnzEti(1) As String
    Dim iBestellt   As Integer
    
    anzeige "normal", "", Label8(2)
    
    If gbAutoZwsp Then
        MSFlexGrid1_SelChange
        Zwischenspeichern_neu
        
        Screen.MousePointer = 11
        neuFildatschreiben
        Screen.MousePointer = 0
    End If
    
    ctmp = Trim(Text5.Text)
    llaeng = Len(ctmp)
    
    If ctmp = "" Then
        ctmp = Trim$(Text7.Text)
        If ctmp = "" Then Exit Sub

        sSQL = "select * from artlief where Trim(libesnr) = '" & ctmp & "' "
        sSQL = sSQL & " and artlief.linr = " & cAnfuLinr & " "
        
        Set rsArt = gdBase.OpenRecordset(sSQL)
        If Not rsArt.EOF Then
            rsArt.MoveFirst
            sArtnr = rsArt!artnr
        End If
        rsArt.Close
    Else
        If llaeng < 7 Then
            sArtnr = ctmp
        ElseIf llaeng >= 7 Then
            
            If Ist_in_ARTEAN_K(ctmp) Then
                
            End If
            
            sSQL = "select artikel.artnr from artikel "
            sSQL = sSQL & " inner join Artlief on artikel.artnr = artlief.artnr "
            
            sSQL = sSQL & " where (ean = '" & ctmp & "'"
            sSQL = sSQL & " or ean2 = '" & ctmp & "'"
            sSQL = sSQL & " or ean3 = '" & ctmp & "')"
            sSQL = sSQL & " and artlief.linr = " & cAnfuLinr & " "
            
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                sArtnr = rsArt!artnr
            End If
            rsArt.Close: Set rsArt = Nothing
        End If
    End If
    
    bFound = False
    
    If sArtnr = "" Then
        Text7.Text = ""
        Text5.Text = ""
        Text5.SetFocus
        
        If Check1.Value = vbUnchecked Then
            Text10.Text = 1
        End If
    Else
        Text5.Text = ""
        lRows = MSFlexGrid1.Rows
        MSFlexGrid1.Redraw = False
        For i = 2 To lRows - 1
            MSFlexGrid1.Row = i
            MSFlexGrid1.Col = 0
            
            
            If MSFlexGrid1.Text = sArtnr Then
                bFound = True
                MSFlexGrid1.TopRow = i
                MSFlexGrid1.Col = 5
                MSFlexGrid1.Row = i
                
                If Check2.Value = vbChecked Then
                    MSFlexGrid1.Text = Val(MSFlexGrid1.Text) + Val(Text10.Text)
                    
                    MSFlexGrid1.Col = 6
                    MSFlexGrid1.Row = i

                    MSFlexGrid1.Text = Val(MSFlexGrid1.Text) + Val(Text10.Text)
                    
                    If isEtidruFree And Check4.Value = vbChecked Then
                        isEtidruFree = False
                        
                        acArtNr(0) = sArtnr
                        acAnzEti(0) = Val(Text10.Text)
                    
                        Select Case cboStrichEndlos.Text
                            Case "69 x 14 (Var 1)" 'Schmucketikett 69x14 Variante 1
                                DruckeSchmucketikett69x14Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL311a", gcEtikettenDrucker, False
                                           
                            Case "69 x 14 (Var 2)"  'Schmucketikett 69x14 Variante 2
                                DruckeSchmucketikett69x14Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL311b", gcEtikettenDrucker, False
                                
                            Case "40 x 18 (Var 1)"  'Etikett 40x18 Variante 1
                                DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL312a", gcEtikettenDrucker, False
                            
                            Case "40 x 18 (Var 2)"  'Etikett 40x18 Variante 2
                                DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL312b", gcEtikettenDrucker, False
                                
                            Case "40 x 18 (Var 3)"  'Etikett 40x18 Variante 3
                                DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL312c", gcEtikettenDrucker, False
                                
                            Case "40 x 18 (Var 4)"  'Etikett 40x18 Variante 4
                                DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL312d", gcEtikettenDrucker, False
                                
                            Case "45 x 23 (Var 1)"  'Etikett 45x23 Variante 1
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL313a", gcEtikettenDrucker, False
                                
                            Case "45 x 23 (Var 2)"  'Etikett 45x23 Variante 2
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL313b", gcEtikettenDrucker, False
                                
                            Case "69 x 14 (Var 3)"  'Schmucketikett 69x14 Variante 3
                                DruckeSchmucketikett69x14Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL311c", gcEtikettenDrucker, False
                                
                            Case "45 x 23 (Var 3)"  'Etikett 45x23 Variante 3
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL313c", gcEtikettenDrucker, False
                                
                            Case "38 x 23 (Var 1)"  'Etikett 38x23 Variante 1
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL314a", gcEtikettenDrucker, False
                                
                            Case "38 x 23 (Var 2)"  'Etikett 38x23 Variante 2
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL314b", gcEtikettenDrucker, False
                                
                            Case "38 x 23 (Var 3)"  'Etikett 38x23 Variante 3
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL314c", gcEtikettenDrucker, False
                                
                            Case "51 x 19 (Var 1)"  'Etikett 51x19 Variante 1
                                DruckeEtikett51x19Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL315a", gcEtikettenDrucker, False
                                
                            Case "51 x 19 (Var 2)"  'Etikett 51x19 Variante 2
                                DruckeEtikett51x19Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL315b", gcEtikettenDrucker, False
                                
                            Case "49 x 19 (Var 1)"  'Etikett 49x19 Variante 1
                                DruckeEtikett49x19Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL316a", gcEtikettenDrucker, False
                                
                            Case "44 x 21 (Var 1)"  'Etikett 44x21 Variante 1
                                DruckeEtikett44x21Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL317a", gcEtikettenDrucker, False
                                
                            Case "51 x 19 (Var 3)"  'Etikett 51x19 Variante 3
                                DruckeEtikett51x19Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL315c", gcEtikettenDrucker, False
                                
                            Case "30 x 15 (Var 1)"  'Etikett 30x15 Variante 1
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL3015a", gcEtikettenDrucker, False
                                
                            Case "30 x 15 (Var 2)"  'Etikett 30x15 Variante 2
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL3015b", gcEtikettenDrucker, False
                                
                            Case "30 x 15 (Var 3)"  'Etikett 30x15 Variante 3
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL3015c", gcEtikettenDrucker, False
                                
                            Case "48 x 18 (Var 1)"  'Etikett 48x18 Variante 1
                                DruckeEtikett48x18Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL319a", gcEtikettenDrucker, False
                                
                            Case "45 x 23 (Var 4)"  'Etikett 45x23 Variante 4
                                DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL313d", gcEtikettenDrucker, False
                               
                            Case "40 x 18 (Var 5)"  'Etikett 40x18 Variante 5
                                DruckeEtikett40x18Variante5 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL312e", gcEtikettenDrucker, False
                                
                            Case "40 x 18 (Var 6)"  'Etikett 40x18 Variante 6
                                DruckeEtikett40x18Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL312f", gcEtikettenDrucker, False
                                
                            Case "35 x 15 (Var 1)" 'Etikett 35x15 Variante 1
                                DruckeEtikett35x15Variante1 acArtNr(), 0, acAnzEti()
                                reportbildschirmToPrinterETI "aWKL322a", gcEtikettenDrucker, False
                                
                        End Select




















                        
                        anzeige "erfolg", "", Label5
                        isEtidruFree = True
                    End If
                    
                    AktualisiereEingangWK15a
                    MoveBestell2GridWK15a cSort
                    
                End If
                
                'Geliefert = Bestellt
                If Check6.Value = vbChecked Then
                    iBestellt = 0
                    MSFlexGrid1.Col = 4
                    MSFlexGrid1.Row = i
                    iBestellt = Val(MSFlexGrid1.Text)
                    
                    MSFlexGrid1.Col = 5
                    MSFlexGrid1.Row = i

                    MSFlexGrid1.Text = iBestellt
                    
                    MSFlexGrid1.Col = 6
                    MSFlexGrid1.Row = i

                    MSFlexGrid1.Text = iBestellt
                    
                    If isEtidruFree And Check4.Value = vbChecked Then
                        isEtidruFree = False
                        
                        acArtNr(0) = sArtnr
                        acAnzEti(0) = iBestellt
                        
                        DruckeEtikett45x23Variante1 acArtNr(), 0, acAnzEti()
                        reportbildschirmToPrinterETI "aWKL3015a", gcEtikettenDrucker, True

                        
                        anzeige "erfolg", "", Label5
                        isEtidruFree = True
                    End If
                    
                    
                    

                    AktualisiereEingangWK15a
                    MoveBestell2GridWK15a cSort
                    
                End If
                
                Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
                Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
                
'                MSFlexGrid1.TopRow = i
'                MSFlexGrid1.Col = 5
'                MSFlexGrid1.Row = i
                
                Text5.Text = ""
                Text7.Text = ""
                
                If Check2.Value = vbUnchecked Then
                    MSFlexGrid1.SetFocus
                Else
                    Text5.SetFocus
                End If
                
                If Check6.Value = vbChecked Then
                    Text5.SetFocus
                End If
                
                Exit For
            Else
                MSFlexGrid1.TopRow = 1
                MSFlexGrid1.Row = 1
            End If
        Next i
        MSFlexGrid1.Redraw = True
        
        
        
        
        
        lRows = MSFlexGrid1.Rows
        MSFlexGrid1.Redraw = False
        For i = 2 To lRows - 1
            MSFlexGrid1.Row = i
            MSFlexGrid1.Col = 0
            
            
            If MSFlexGrid1.Text = sArtnr Then
                bFound = True
                MSFlexGrid1.TopRow = i
                MSFlexGrid1.Col = 5
                MSFlexGrid1.Row = i
                Exit For
            End If
        Next i
        MSFlexGrid1.Redraw = True
    
    End If
    
    If bFound = False Then
        If sArtnr <> "" Then
            sBez = bezis(sArtnr)
            If sBez <> "" Then
                sSQL = "Select * from  " & sZufall
                Set rsArt = gdBase.OpenRecordset(sSQL)
    
                rsArt.AddNew
                rsArt!artnr = sArtnr
                rsArt!BEZEICH = sBez
                rsArt!lekpr = ermLEKPR(sArtnr, CLng(cAnfuLinr))
                rsArt!BESTELLT = 0
                
                If Check2.Value = vbChecked Then
                    rsArt!GELIEFERT = 1
                    rsArt!BERECHNET = 1
                Else
                    rsArt!GELIEFERT = 0
                    rsArt!BERECHNET = 0
                End If
                
                rsArt!LIEFBETRAG = 0
                rsArt!ZEILEN_RAB = 0
                rsArt!ZEILENWERT = 0
                rsArt!RECHN_RAB = 0
                rsArt!RECHN_WERT = 0
                rsArt!STCK_PREIS = 0
                rsArt!linr = cAnfuLinr
                rsArt!LIBESNR = 0
                rsArt!KVKPR1 = ermKVKPR1(sArtnr)
                
                
                Text9.Text = Format(rsArt!lekpr, "#####0.00")
                    
                rsArt.Update
                rsArt.Close: Set rsArt = Nothing
                
                MoveBestell2GridWK15a cSort
                
               
                Text9.SetFocus
                
                anzeige "laser", "Dieser Artikel(" & sArtnr & ") wurde gerade angefügt.", Label8(2)
            End If
        End If
    End If
    
    If sArtnr <> "" Then
        Label3(9).Caption = sArtnr
        Label3(11).Caption = "(" & ermBESTAND(sArtnr) & ")"
    End If
    
    If Check1.Value = vbUnchecked Then
        Text10.Text = 1
    End If
    
    Text7.Text = ""
    Text5.Text = ""
'    Text5.SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command7_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
    
        Case 0
            del_DESADV
        
        Case 1
            Frame4.Visible = False
        Case 2
            
            DESADV_anwenden
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DESADV_anwenden()
On Error GoTo LOKAL_ERROR

    Dim sAuftragsnr     As String
    

    If List5.ListIndex < 0 Then
        MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
        List5.SetFocus
    Else
        sAuftragsnr = List5.list(List5.ListIndex)
        sAuftragsnr = Trim(Mid$(sAuftragsnr, 1, InStr(1, sAuftragsnr, " ")))
        
        If Left(Label3(10).Caption, 6) = sAuftragsnr Then
        
            DESADV_anwenden_Einzel sAuftragsnr
            Frame4.Visible = False
        
        Else
            MsgBox "Bitte den richtigen Lieferschein auswählen. Erwartet wird: " & Left(Label3(10).Caption, 6) & "", vbInformation, "Winkiss Hinweis:"
            
            setzefokus Label3(10).Caption
            
'            List5.SetFocus
        End If
    End If

   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DESADV_anwenden"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub setzefokus(sAnr As String)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    
    For i = 0 To List5.ListCount - 1
        List5.Selected(i) = False
    Next i

    For i = 0 To List5.ListCount - 1
        If Trim(List5.list(i)) = sAnr Then
            List5.Selected(i) = True
        End If
    Next i
        
    
    
    List5.SetFocus
    
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "setzefokus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DESADV_anwenden_Einzel(sANummer As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim sArtnr      As String
    Dim rsArt       As DAO.Recordset
    Dim rsrs        As DAO.Recordset
    Dim sMenge      As String
    
    Screen.MousePointer = 11
    
    '1. alles - geliefert auf 0 setzen
    sSQL = "Update " & sZufall & " Set GELIEFERT = 0, BERECHNET = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    '2. Menge hinzufügen
    sSQL = "select * from DESADV where AUFTRAGSnr = " & sANummer & ""
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
    
        rsArt.MoveFirst
        Do While Not rsArt.EOF
        
            sArtnr = 0
            If Not IsNull(rsArt!artnr) Then
                sArtnr = rsArt!artnr
            End If
            
'            If sArtnr = 610338 Then
'                MsgBox "Hey"
'            End If
'
            sMenge = 0
            If Not IsNull(rsArt!Menge) Then
                sMenge = rsArt!Menge
            End If
            
            
            If sArtnr > 0 Then
            
                If DatendrinSQL("Select * from " & sZufall & "  where artnr = " & sArtnr, gdBase) Then
                    
                    sSQL = "Update " & sZufall
                    sSQL = sSQL & " Set GELIEFERT  = GELIEFERT + " & sMenge
                    sSQL = sSQL & " , BERECHNET  = BERECHNET + " & sMenge
                    sSQL = sSQL & " where artnr = " & sArtnr
                    gdBase.Execute sSQL, dbFailOnError
                
                Else
            
                    If DatendrinSQL("Select * from Artlief where artnr = " & sArtnr & " and Linr = " & cAnfuLinr, gdBase) Then
                    
                        sSQL = "Select * from " & sZufall & " where ARTNR = " & sArtnr
                        Set rsrs = gdBase.OpenRecordset(sSQL)
                        If rsrs.EOF Then
                            rsrs.AddNew
                            rsrs!artnr = sArtnr
                            rsrs!BEZEICH = bezis(sArtnr)
                            rsrs!lekpr = ermLEKPR(sArtnr, CLng(cAnfuLinr))
                            rsrs!BESTELLT = 0
                            rsrs!GELIEFERT = sMenge
                            rsrs!BERECHNET = sMenge
                            rsrs!LIEFBETRAG = 0
                            rsrs!ZEILEN_RAB = 0
                            rsrs!ZEILENWERT = 0
                            rsrs!RECHN_RAB = 0
                            rsrs!RECHN_WERT = 0
                            rsrs!STCK_PREIS = 0
                            rsrs!linr = cAnfuLinr
                            rsrs!LIBESNR = 0
                            rsrs!KVKPR1 = ermKVKPR(sArtnr)
            '                                rsRs!AWM = "99"
                            rsrs.Update
                        End If
                        rsrs.Close
                    End If
                End If

            End If
            
            rsArt.MoveNext
        Loop
        
    End If
    rsArt.Close: Set rsArt = Nothing
    
    FormatiereGridWK15a
    MoveBestell2GridWK15a cSort

    Screen.MousePointer = 0
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DESADV_anwenden_Einzel"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub del_DESADV()
On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim iRet            As Integer
    Dim sAuftragsnr     As String

    If List5.ListIndex < 0 Then
        iRet = MsgBox("Möchten Sie alle Dateien löschen?", vbYesNo + vbQuestion + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            For i = 0 To List5.ListCount - 1
                sAuftragsnr = List5.list(i)
                sAuftragsnr = Trim(Mid$(sAuftragsnr, 1, InStr(1, sAuftragsnr, " ")))
                
                del_DESADV_Einzel sAuftragsnr
            Next i
        End If
    Else
        sAuftragsnr = List5.list(List5.ListIndex)
        sAuftragsnr = Trim(Mid$(sAuftragsnr, 1, InStr(1, sAuftragsnr, " ")))
        
        del_DESADV_Einzel sAuftragsnr
    End If

    fuelle_Frame4_mit_DESADV
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "del_DESADV"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub del_DESADV_Einzel(sANummer As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
   
    sSQL = "Delete from DESADV where AUFTRAGSNR = " & sANummer
    gdBase.Execute sSQL, dbFailOnError
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "del_DESADV_Einzel"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    gbGescheitert = False
    
    Screen.MousePointer = 11

    PositionierenWK15a
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Frame4.BackColor = glH2
    Label3(8).BackColor = glH2
    
    
    fraArtAnfuegen.Visible = False
    Text1.Text = ""
    Text5.Text = ""
    Text7.Text = ""
    
    gbAender = False
    gbUpdate = False
    
    If gbDELBDAT Then
        Check17.Value = vbChecked
    Else
        Check17.Value = vbUnchecked
    End If
    
    If gbDIFFPROT Then
        Check3.Value = vbChecked
    Else
        Check3.Value = vbUnchecked
    End If
    
    If gbUEBERPROT Then
        Check5.Value = vbChecked
    Else
        Check5.Value = vbUnchecked
    End If
    
    fülleCboEtiketten cboStrichEndlos
    
    
    
    If NewTableSuchenDBKombi("E15A", gdApp) Then
    
        If SpalteInTabellegefundenNEW("E15A", "Eti", gdApp) = False Then
            SpalteAnfuegenNEW "E15A", "Eti", "Text(20)", gdApp
        End If
        voreinstellungladen
    End If
    
    If NewTableSuchenDBKombi("LASTFOCUS", gdBase) = False Then
        CreateTableT2 "LASTFOCUS", gdBase
    End If
    
    
    
    neuFildatschreiben
    
    LeseInhaltWK15a
    
    isEtidruFree = True
    
    If Option1(0).Value = True Then
        cSort = " order by MOPREIS, BEZEICH"
    ElseIf Option1(1).Value = True Then
        cSort = " order by MOPREIS, val(LIBESNR) asc "
    ElseIf Option1(2).Value = True Then
        cSort = " order by MOPREIS, ARTNR"
    ElseIf Option1(3).Value = True Then
        cSort = " order by MOPREIS, LPZ"
    Else
        cSort = " order by MOPREIS, LPZ"
    End If
    
    Dim i As Integer
        
    For i = 12 To 15
        Command0(i).BackColor = vbWhite
        Command0(i).HoverColorFrom = vbWhite
        Command0(i).HoverColorTo = vbWhite
    Next i
    
    
    
    
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungladen()
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim sEti As String
    Set rsrs = gdApp.OpenRecordset("E15A")
    
    If Not rsrs.EOF Then
        
        Option1(0).Value = rsrs!bo1
        Option1(1).Value = rsrs!bo2
        Option1(2).Value = rsrs!bo3
        Option1(3).Value = rsrs!bo4
        
        Option2(0).Value = rsrs!bo5
        Option2(1).Value = rsrs!bo6
        Option2(2).Value = rsrs!bo7
        
        sEti = "bitte auswählen"
        If Not IsNull(rsrs!Eti) Then
            sEti = Trim(rsrs!Eti)
        End If

        
        cboStrichEndlos.Text = sEti
        
    
        If rsrs!bo8 = True Then
            Check2.Value = vbUnchecked
        Else
            Check2.Value = vbChecked
        End If
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer
    Dim bo6 As Integer
    Dim bo7 As Integer
    Dim bo8 As Integer
    Dim bo9 As Integer
    
    loeschNEW "E15A", gdApp
    CreateTable "E15A", gdApp
    
    bo1 = Option1(0).Value
    bo2 = Option1(1).Value
    bo3 = Option1(2).Value
    bo4 = Option1(3).Value
    bo5 = Option2(0).Value
    bo6 = Option2(1).Value
    bo7 = Option2(2).Value
    
    If Check2.Value = vbChecked Then
        bo8 = 0
    Else
        bo8 = -1
    End If
        
    sSQL = "Insert into E15A (BO1,BO2,BO3,BO4,BO5,BO6,BO7,BO8,Eti) "
    sSQL = sSQL & " values (" & bo1 & "," & bo2 & "," & bo3 & "," & bo4 & "," & bo5 & "," & bo6 & "," & bo7 & "," & bo8 & ",'" & cboStrichEndlos.Text & "'"
    sSQL = sSQL & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR

    LogtoEnd Me
    voreinstellungspeichern
    
    loeschNEW "BESTIMP", gdBase
    loeschNEW "DRU_TEMP", gdBase
    loeschNEW "DT" & srechnertab, gdBase
    
    If sZufall <> "" Then
        loesch sZufall
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Label3(6).ForeColor = glS1

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame2_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Image2_Click()
    On Error GoTo LOKAL_ERROR
    
    MdeVerarbeitung
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Image2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WeiterverarbeitungMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim RsMDEIN     As Recordset
    Dim rsArt       As Recordset
    Dim rsrs        As Recordset
    Dim sArtnr      As String
    Dim sEAN        As String
    Dim sMenge      As String
    
    Set RsMDEIN = gdBase.OpenRecordset("mdein", dbOpenTable)
    If Not RsMDEIN.EOF Then
        RsMDEIN.MoveFirst
        Do While Not RsMDEIN.EOF
        
            If Not IsNull(RsMDEIN!eancode) Then
                sEAN = Trim(RsMDEIN!eancode)
                sEAN = checkean(sEAN)
            Else
                sEAN = ""
            End If
'            sEAN = IIf(IsNull(RsMDEIN(0)), "", RsMDEIN(0))
'            sEAN = checkean(sEAN)
            
            If sEAN <> "" Then
                If IsNumeric(sEAN) Then
                    If Len(sEAN) = 11 Then
                        sEAN = "0" & sEAN
                
                        cSQL = "select * from artikel where ean = '" & sEAN & "'"
                        cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                        cSQL = cSQL & " or ean3 = '" & sEAN & "'"
                    ElseIf Len(sEAN) = 8 Then
                    
                        If Left(sEAN, 1) = "2" Then
                            sEAN = Mid$(sEAN, 2, 6)
                            cSQL = "select * from artikel where artnr = " & sEAN
                        Else
                            cSQL = "select * from artikel where ean = '" & sEAN & "'"
                            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
                        End If
                    ElseIf Len(sEAN) = 6 Then
                        cSQL = "select * from artikel where artnr = " & sEAN
                        
                    Else
                        cSQL = "select * from artikel where ean = '" & sEAN & "'"
                        cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                        cSQL = cSQL & " or ean3 = '" & sEAN & "'"
                    End If
                    
                    Set rsArt = gdBase.OpenRecordset(cSQL)
                    If Not rsArt.EOF Then
                        sArtnr = Trim(rsArt!artnr)
                    Else
                        sArtnr = 0
                    End If
                    rsArt.Close: Set rsArt = Nothing
                    
                    sMenge = IIf(IsNull(RsMDEIN(1)), "", RsMDEIN(1))
                    
                    If DatendrinSQL("Select * from " & sZufall & "  where artnr = " & sArtnr, gdBase) Then
                    
                        cSQL = "Update " & sZufall
                        cSQL = cSQL & " Set GELIEFERT  = GELIEFERT + " & sMenge
                        cSQL = cSQL & " , BERECHNET  = BERECHNET + " & sMenge
                        cSQL = cSQL & " where artnr = " & sArtnr
                        gdBase.Execute cSQL, dbFailOnError
                    
                    Else
            
                        If DatendrinSQL("Select * from Artlief where artnr = " & sArtnr & " and Linr = " & cAnfuLinr, gdBase) Then
                        
                            cSQL = "Select * from " & sZufall & " where ARTNR = " & sArtnr
                            Set rsrs = gdBase.OpenRecordset(cSQL)
                            If rsrs.EOF Then
                                rsrs.AddNew
                                rsrs!artnr = sArtnr
                                rsrs!BEZEICH = bezis(sArtnr)
                                rsrs!lekpr = ermLEKPR(sArtnr, CLng(cAnfuLinr))
                                rsrs!BESTELLT = 0
                                rsrs!GELIEFERT = sMenge
                                rsrs!BERECHNET = sMenge
                                rsrs!LIEFBETRAG = 0
                                rsrs!ZEILEN_RAB = 0
                                rsrs!ZEILENWERT = 0
                                rsrs!RECHN_RAB = 0
                                rsrs!RECHN_WERT = 0
                                rsrs!STCK_PREIS = 0
                                rsrs!linr = cAnfuLinr
                                rsrs!LIBESNR = 0
                                rsrs!KVKPR1 = ermKVKPR(sArtnr)
'                                rsRs!AWM = "99"
                                rsrs.Update
                            End If
                            rsrs.Close
                        End If
                    End If
                End If
            End If
            
            RsMDEIN.MoveNext
        Loop
    End If
    RsMDEIN.Close
    
    FormatiereGridWK15a
    MoveBestell2GridWK15a cSort
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WeiterverarbeitungMDE"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub MdeVerarbeitung()
    On Error GoTo LOKAL_ERROR

        Dim sSQL As String
        
        Screen.MousePointer = 11
        
        If MDEeinlesenOhneLinr(lblUeberschrift(1), txtStatus, picprogress, frmWK15a) = True Then
        
            loeschNEW "MDEIN", gdBase
        
            sSQL = "Select EANCODE, sum(Menge)as Meng INTO MDEIN FROM mdeinh group by EANCODE "
            gdBase.Execute sSQL, dbFailOnError
    
'            Command5_Click 0
            WeiterverarbeitungMDE
            
        End If
        Screen.MousePointer = 0
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MdeVerarbeitung"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub fuelle_Frame4_mit_DESADV()
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsrs    As DAO.Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    
    cSQL = "Select distinct(auftragsnr) as aufnr from DESADV order by auftragsnr desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    List5.Clear
    
    If Not rsrs.EOF Then
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!aufnr) Then
                cFeld = rsrs!aufnr
            End If
    
            cLBSatz = cFeld & Space(7 - Len(cFeld))
            
            List5.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    setzefokus Label3(10).Caption

    List5.Refresh
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelle_Frame4_mit_DESADV"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Label1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lMDHDAT         As Long
    Dim lKJADate        As Long
    Dim cKJAZeit        As String
    
    If Index = 2 Then
        Frame4.Visible = True
        fuelle_Frame4_mit_DESADV
        
    End If
    
    
    

    If Index = 1 Then
    
        lKJADate = Fix(Now)
        cKJAZeit = Format$(Now, "HH:MM:SS")
        
        If IsDate(Text8.Text) = False Then
            MsgBox "Bitte geben Sie ein Datum an! ('TT.MM.JJ')", vbInformation, "Winkiss Hinweis:"
            Text8.SetFocus
            Exit Sub
        End If
        
        If Label3(9).Caption = "" Then
            MsgBox "Bitte markieren Sie eine Zeile", vbInformation, "Winkiss Hinweis:"
            Exit Sub
        End If
        
        'speichern
        If Text8.Text <> "" Then
            lMDHDAT = DateValue(Text8.Text)
            insertArtikelMDH lKJADate, cKJAZeit, CInt(gcBedienerNr), CLng(Label3(9).Caption), lMDHDAT
            Label3(9).Caption = ""
            
            Text5.SetFocus 'Fokus auf Artikelsuche
            
        End If
        
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label2_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 19 Then
        frmWKL192.Show 1
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label3_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 6
            gsLinr = Label3(6).Caption
            frmWKL17.Show 1
            gsLinr = ""
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 6
        Label3(6).ForeColor = glLink
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label3_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_Click()
    On Error GoTo LOKAL_ERROR
    
    glSelect = MSFlexGrid1.Row
    Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
    
    MSFlexGrid1.Col = 12
    Text9.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 0
    Label0(2).Caption = MSFlexGrid1.Text
    Label3(9).Caption = MSFlexGrid1.Text 'Artnr für MDH
    Label3(11).Caption = "(" & ermBESTAND(Label3(9).Caption) & ")"
    MSFlexGrid1.Col = Val(Label0(1).Caption)

    giErsetzen = 0
   
    If gbUpdate Then
        MoveBestell2GridWK15a cSort
        gbUpdate = False
    End If
    
    If glSelect = glmaxtabzeilenanz + 1 Then
        MSFlexGrid1.TopRow = glSelect
        MSFlexGrid1.Row = glSelect
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid1.Row > 1 Then
        
    Else
        If MSFlexGrid1.Col = 2 Then 'Bezeichnung
            If byteSortReihen = 1 Then
                cSort = " order by MOPREIS, BEZEICH desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by MOPREIS, BEZEICH asc"
            End If
            sortierenHGrid MSFlexGrid1
        ElseIf MSFlexGrid1.Col = 1 Then 'Libesnr
            If byteSortReihen = 1 Then
                cSort = " order by MOPREIS, val(LIBESNR) desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by MOPREIS, val(LIBESNR) asc"
            End If
            sortierenHGrid MSFlexGrid1
        ElseIf MSFlexGrid1.Col = 3 Then 'Lekpr
            If byteSortReihen = 1 Then
                cSort = " order by lekpr desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by lekpr asc"
            End If
            sortierenHGrid MSFlexGrid1
        ElseIf MSFlexGrid1.Col = 0 Then 'artnr
            If byteSortReihen = 1 Then
                cSort = " order by MOPREIS, ARTNR desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by MOPREIS, ARTNR asc"
            End If
            sortierenHGrid MSFlexGrid1
        ElseIf MSFlexGrid1.Col = 16 Then 'Linie
            If byteSortReihen = 1 Then
                cSort = " order by MOPREIS, LPZ desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by MOPREIS, LPZ asc"
            End If
            sortierenHGrid MSFlexGrid1
        ElseIf MSFlexGrid1.Col = 4 Then 'bestellt
            If byteSortReihen = 1 Then
                cSort = " order by BESTELLT desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by BESTELLT asc"
            End If
            sortierenHGrid MSFlexGrid1
            
        ElseIf MSFlexGrid1.Col = 5 Then 'GELIEFERT
            If byteSortReihen = 1 Then
                cSort = " order by GELIEFERT desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by GELIEFERT asc"
            End If
            sortierenHGrid MSFlexGrid1
        ElseIf MSFlexGrid1.Col = 6 Then 'BERECHNET
            If byteSortReihen = 1 Then
                cSort = " order by BERECHNET desc"
            ElseIf byteSortReihen = 2 Then
                cSort = " order by BERECHNET asc"
            End If
            sortierenHGrid MSFlexGrid1
        End If
        
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen    As String
    Dim cFeld       As String
    Dim lcol        As Long
    Dim lrow        As Long
    Dim cValid      As String
    
    lrow = Label0(0).Caption
    lcol = Label0(1).Caption
    
    'hier mal kurz Zwischenspeichern
    If gbAutoZwsp Then
        Zwischenspeichern
        
                
        Screen.MousePointer = 11
        neuFildatschreiben
        Screen.MousePointer = 0
    End If
    'Ende hier mal kurz Zwischenspeichern
    
    '***********************************************
    '* Erste und letzte Zeile sind nicht edit-fähig
    '***********************************************
    If lrow = 0 Then 'Or lRow = MSFlexGrid1.Rows - 1
        KeyAscii = 0
        Exit Sub
    End If
    
    '***********************************************
    '* Bestimmte Spalten sind nicht edit-fähig
    '***********************************************
    If lcol < 3 Or lcol = 7 Or lcol = 9 Or lcol = 11 Or lcol = 12 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    '***********************************************
    '* Für den Rest gilt:
    '***********************************************

    Select Case lcol
        Case 3, 8, 10, 13
            cValid = "1234567890," & Chr$(8)
        Case 4 To 6
            cValid = "1234567890" & Chr$(8)
    End Select
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    If giErsetzen > 0 Then
        cFeld = MSFlexGrid1.TextMatrix(lrow, lcol)
    Else
        cFeld = ""
    End If
    
    Select Case KeyAscii
        Case Is = 8
            If Len(cFeld) > 0 Then
                cFeld = Left(cFeld, Len(cFeld) - 1)
            End If
        
        Case Else
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            If cZeichen = "," Then
                If InStr(cFeld, ",") = 0 Then
                    cFeld = cFeld & Chr$(KeyAscii)
                End If
            Else
                cFeld = cFeld & Chr$(KeyAscii)
            End If
    End Select
    
    MSFlexGrid1.TextMatrix(lrow, lcol) = cFeld
    
    gbAender = True
    giErsetzen = 2
    
    MSFlexGrid1.Row = lrow
    MSFlexGrid1.Col = lcol
    MSFlexGrid1.SetFocus
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSFlexGrid1.Row
    lcol = MSFlexGrid1.Col
    
    Select Case KeyCode
        Case Is = 46    'Del
            MSFlexGrid1.TextMatrix(lrow, lcol) = ""
            
            gbAender = True
            gbUpdate = True
            
        Case Is = vbKeyF2
            lrow = MSFlexGrid1.Row
            gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
            If gsARTNR <> "" Then
    
                frmWKL10.Show 1
                Me.Refresh
                Screen.MousePointer = 11
                MSFlexGrid1.Col = 5
                MSFlexGrid1.Row = lrow
                MSFlexGrid1.TopRow = lrow
                MSFlexGrid1.SetFocus
                Screen.MousePointer = 0
            End If
            gsARTNR = ""
        Case Is = vbKeyReturn
            Text5.SetFocus
        
    End Select
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub UpdateEingangWKL10(cArtNr As String)
    On Error GoTo LOKAL_ERROR
            
    Dim lcol        As Long
    Dim lcount      As Long
    Dim cFeldName   As String
    Dim cWert       As String
    
    Dim cSQL        As String
    Dim dWert       As Double
    Dim iFeldType   As Integer
    Dim bRechnRab   As Boolean
    Dim rsrs        As Recordset
    Dim iRet        As Integer
    
    MSFlexGrid1.Redraw = False
    bRechnRab = False
    
    If Label0(2).Caption = "" Then
        Exit Sub
    End If

    lcol = Label0(1).Caption
    cWert = Label0(3).Caption
    
    If lcol = 3 Or lcol = 8 Or lcol = 10 Or lcol = 12 Or lcol = 13 Then
        cWert = fnMoveComma2Point$(cWert)
    End If
    dWert = Val(cWert)
    
    Set rsrs = gdBase.OpenRecordset("Select * from " & Trim(sZufall) & " where artnr = " & cArtNr)
    
    If Not rsrs.EOF Then
        rsrs.Edit
        Select Case lcol
            Case Is = 3
                rsrs!lekpr = dWert
                
                Artikelveraenderung cArtNr, CStr(dWert), "WE aus Best", "LEKPR", cAnfuLinr

            Case Is = 4
            
                If dWert > 2000 Then 'Frage 301106
                    iRet = MsgBox("Der Wert liegt bei " & dWert & ". Wirklich?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbNo Then
                        
                        Exit Sub
                    End If
                End If
                
                If dWert > 30000 Then
                    Exit Sub
                End If
                
                rsrs!BESTELLT = dWert
            Case Is = 5
            
                If dWert > 9999 Then 'Frage 301106
                    iRet = MsgBox("Der Wert liegt bei " & dWert & ". Wirklich?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbNo Then
                        
                        Exit Sub
                    End If
                End If
                
                If dWert > 30000 Then
                    Exit Sub
                End If
                
                rsrs!GELIEFERT = dWert
                
                rsrs!BERECHNET = dWert
                
                If IsNull(rsrs!GELIEFERT) Or rsrs!GELIEFERT = 0 Then
                    rsrs!STCK_PREIS = rsrs!lekpr
                Else
                    rsrs!STCK_PREIS = ((rsrs!lekpr * rsrs!BERECHNET) / rsrs!GELIEFERT)
                End If
                
                
            Case Is = 6 'BERECHNET 301106
                If dWert > 2000 Then 'Frage
                    iRet = MsgBox("Der Wert liegt bei " & dWert & ". Wirklich?", vbYesNo + vbDefaultButton2 + vbQuestion, "Winkiss Frage:")
                    If iRet = vbNo Then
                        
                        Exit Sub
                    End If
                End If
                
                If dWert > 30000 Then
                    Exit Sub
                End If
                
                rsrs!BERECHNET = dWert
                
                If IsNull(rsrs!GELIEFERT) Or rsrs!GELIEFERT = 0 Then
                    rsrs!STCK_PREIS = rsrs!lekpr
                Else
                    rsrs!STCK_PREIS = ((rsrs!lekpr * rsrs!BERECHNET) / rsrs!GELIEFERT)
                End If
            Case Is = 7
                rsrs!LIEFBETRAG = dWert
            Case Is = 8
                rsrs!ZEILEN_RAB = dWert
            Case Is = 9
                rsrs!ZEILENWERT = dWert
            Case Is = 10
                rsrs!RECHN_RAB = dWert
                bRechnRab = True
            Case Is = 11
                rsrs!RECHN_WERT = dWert
            Case Is = 12
                rsrs!STCK_PREIS = dWert
                
            Case Is = 13
                rsrs!KVKPR1 = dWert
                
                Artikelveraenderung cArtNr, CStr(dWert), "WE aus Best", "KVKPR1"
        End Select
        rsrs.Update
        
        If bRechnRab Then
            cSQL = "Update " & sZufall & "  set RECHN_RAB = " & Trim$(Str$(dWert))
            gdBase.Execute cSQL, dbFailOnError
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    gbUpdate = True
    gbAender = False

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateEingangWKL10"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_SelChange()
    On Error GoTo LOKAL_ERROR
    
    Dim lColmerker As Long
    Dim lRowmerker As Long
    
    '**********************************************
    '* Diesen ganzen Quatsch irgendwann umstellen
    '* auf die TextMatrix-Funktion!
    '**********************************************
    If gbAender Then
        '** Zielzelle merken **
        lRowmerker = MSFlexGrid1.Row
        lColmerker = MSFlexGrid1.Col
        '** Artikelnummer holen **
        MSFlexGrid1.Row = Val(Label0(0).Caption)
        MSFlexGrid1.Col = 0
        Label0(2).Caption = MSFlexGrid1.Text
        If Label0(4).Caption = "" Then
            Label0(4).Caption = Trim$(Str$(lRowmerker))
            Label0(5).Caption = Trim$(Str$(lColmerker))
        End If
        '** Auf Quellzelle zurücksetzen **
        MSFlexGrid1.Row = Val(Label0(0).Caption)
        MSFlexGrid1.Col = Val(Label0(1).Caption)
        '** Wert der Quellzelle (zu speichernder Wert) holen **
        Label0(3).Caption = MSFlexGrid1.Text
        UpdateEingangWKL10 Label0(2).Caption
        MSFlexGrid1.Row = Val(Label0(4).Caption)
        MSFlexGrid1.Col = Val(Label0(5).Caption)
    End If

    MSFlexGrid1_Click

    If Label0(0).Caption <> Trim$(Str$(MSFlexGrid1.Row)) Then
        Label0(0).Caption = Trim$(Str$(MSFlexGrid1.Row))
        Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))
        MSFlexGrid1.Col = 0
        Label0(2).Caption = MSFlexGrid1.Text
        MSFlexGrid1.Col = Val(Label0(1).Caption)
    End If
    Label0(1).Caption = Trim$(Str$(MSFlexGrid1.Col))

    If Label0(4).Caption <> "" Then
        MSFlexGrid1.Row = Val(Label0(4).Caption)
        MSFlexGrid1.Col = Val(Label0(5).Caption)
        Label0(0).Caption = Label0(4).Caption
        Label0(1).Caption = Label0(5).Caption
        Label0(4).Caption = ""
        Label0(5).Caption = ""
    End If
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
End Sub
Private Sub SaveIndirektWK15a()
    On Error GoTo LOKAL_ERROR
    
    If gbAender Then
        '** Artikelnummer holen **
        MSFlexGrid1.Row = Val(Label0(0).Caption)
        MSFlexGrid1.Col = 0
        Label0(2).Caption = MSFlexGrid1.Text
        
        '** Auf Quellzelle zurücksetzen **
        MSFlexGrid1.Row = Val(Label0(0).Caption)
        MSFlexGrid1.Col = Val(Label0(1).Caption)
        
        '** Wert der Quellzelle (zu speichernder Wert) holen **
        Label0(3).Caption = MSFlexGrid1.Text
        
        UpdateEingangWKL10 Label0(2).Caption

        MSFlexGrid1.Row = Val(Label0(4).Caption)
        MSFlexGrid1.Col = Val(Label0(5).Caption)
    End If

    MSFlexGrid1_Click

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SaveIndirektWK15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Option2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case 0
            ListeFuellAnfangsbuchdataT "Q", List2, "tabname", Label3(3)
        Case 1
            ListeFuellAnfangsbuchdataT "Q", List2, "tabdate", Label3(3)
        Case 2
            ListeFuellAnfangsbuchdataT "Q", List2, "Liefbez", Label3(3)
        Case 3
            ListeFuellAnfangsbuchdataT "Q", List2, "AUFTRAGSNR", Label3(3)
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus()
    On Error GoTo LOKAL_ERROR

    Text1.BackColor = glSelBack1
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Label0(0).Caption = "-999"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Label1_Click 1 'speichern
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text8_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub

Private Sub Text9_GotFocus()
On Error GoTo LOKAL_ERROR

    Text9.BackColor = glSelBack1
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text9_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
    If cZeichen = "," Then
        If InStr(Text9.Text, ",") <> 0 Then
            KeyAscii = 0
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text9_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String

    If KeyCode = vbKeyReturn Then
    
        Dim dLEK As Double
        Dim dREK As Double
        Dim dRabatt As Double
        
        dREK = CDbl(Text9.Text)
        dLEK = ermLEKPR(Label3(9).Caption, Label3(6).Caption)
        
        If dLEK <> 0 Then
            dRabatt = 100 - (dREK * 100 / dLEK)
    
            cSQL = "Update " & sZufall & " Set ZEILEN_RAB = '" & dRabatt & "' where Artnr = " & CLng(Label3(9).Caption)
            gdBase.Execute cSQL, dbFailOnError
        End If

        MoveBestell2GridWK15a cSort
        
        
'        Dim lRows As Long
'        Dim i As Integer
'
'        lRows = MSFlexGrid1.Rows
'        MSFlexGrid1.Redraw = False
'        For i = 2 To lRows - 1
'            MSFlexGrid1.Row = i
'            MSFlexGrid1.Col = 0
'
'
'            If MSFlexGrid1.Text = Label3(9).Caption Then
'
'                MSFlexGrid1.TopRow = i
'                MSFlexGrid1.Col = 5
'                MSFlexGrid1.Row = i
'                MSFlexGrid1.SetFocus
'                Exit For
'            End If
'        Next i
'        MSFlexGrid1.Redraw = True

        
        
        Text5.SetFocus
        
        
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text9_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub

Private Sub Text9_LostFocus()
On Error GoTo LOKAL_ERROR
    Text9.BackColor = vbWhite
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text9_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub txtLieferschein_GotFocus()
    On Error GoTo LOKAL_ERROR

    txtLieferschein.BackColor = glSelBack1
    txtLieferschein.SelStart = 0
    txtLieferschein.SelLength = Len(txtLieferschein.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtLieferschein_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text8_GotFocus()
    On Error GoTo LOKAL_ERROR

    Text8.BackColor = glSelBack1
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8.Text)
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text8_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890," & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If cZeichen = "," Then
        If InStr(Text1.Text, ",") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text10_GotFocus()
On Error GoTo LOKAL_ERROR
    Text10.BackColor = glSelBack1
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text10_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)

    cValid = "1234567890" & Chr$(8)
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text10_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        If Text5.Text <> "" Then
            Command6_Click
        Else
            Text5.SetFocus
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text10_Keyup"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub Text10_LostFocus()
On Error GoTo LOKAL_ERROR
    Text10.BackColor = vbWhite
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text10_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr$(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If cZeichen = "," Then
        If InStr(Text2.Text, ",") > 0 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtLieferschein_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtLieferschein.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtLieferschein_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text8_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text8.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text8_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdAnfuegen_Click 0
    Else
        artikel_suchen
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text4_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text7_GotFocus()
On Error GoTo LOKAL_ERROR

    Text7.BackColor = glSelBack1
    MSFlexGrid1_SelChange
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_GotFocus()
On Error GoTo LOKAL_ERROR

    Text5.BackColor = glSelBack1
    MSFlexGrid1_SelChange
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command6_Click
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command6_Click
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub CheckKundbest(cSuch As String)
On Error GoTo LOKAL_ERROR

    '1. Befindet sich der Artikel unter bestellt
    '2. dann updaten auf nichtgeliefert
    
    Dim cSQL                    As String
    
    cSQL = "Update KUNDBEST set StatusARTIKEL = 'NICHTGELIEFERT' " ',GELIEFERTAM = DateValue(Now)"
    cSQL = cSQL & " where artnr = " & Val(cSuch)
    cSQL = cSQL & " and StatusARTIKEL = 'BESTELLT'"
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "CheckKundbest"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub EingangDerArtikel()
    On Error GoTo LOKAL_ERROR
    
    Dim lBestandAlt     As Long
    Dim lBestandneu     As Long
    Dim lBestandZugang  As Long
    Dim lBestellt       As Long
    Dim lHeute          As Long
    Dim cJetzt          As String
    Dim cSQL            As String
    Dim cArtNr          As String
    Dim cBezeich        As String
    Dim cEAN            As String
    Dim cLinr           As String
    Dim cdatei          As String
    Dim cPfad           As String
    Dim ctmp            As String
    Dim cLiBesNr        As String
    Dim dStückPreis     As Double
    Dim dStückPreisAlt  As Double
    Dim dGesamtPreis    As Double
    Dim dGesamtBestand  As Double
    Dim dGesamtPreisAlt As Double
    Dim dGesamtPreisNeu As Double
    Dim dKVkPr1         As Double
    Dim dAlt            As Double
    Dim dEkpr           As Double
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim rsHis           As Recordset
    Dim rsArtlief       As Recordset
    Dim rsRest          As Recordset
    Dim rsBest          As Recordset
    Dim rsEtiLs         As Recordset
    Dim siAnzeige       As Single
    Dim lcount          As Long
    Dim iStufe          As Integer
    Dim lKJADate        As Long
    Dim cKJAZeit        As String
    
    lKJADate = Fix(Now)
    cKJAZeit = Format$(Now, "HH:MM:SS")
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
        
    '****************************************************************
    '* Die Datei EINGANG auf den Stand der Flexgrid-Anzeige bringen
    '****************************************************************
    iStufe = 1
    AktualisiereEingangWK15a
    
    Screen.MousePointer = 11
    
    cdatei = List2.list(List2.ListIndex)
    cdatei = Trim(Left(cdatei, 10))
    cdatei = UCase$(cdatei)
    
    If gbDELBDAT = True Then 'keine Restbestelldateien
        cSQL = "Select * from " & sZufall & " where geliefert =0"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!artnr) Then
                    'wenn geliefert = 0 dann mal in kundbest gucken
                    CheckKundbest rsrs!artnr
                End If
            rsrs.MoveNext
            Loop
        End If
        rsrs.Close
    End If
    
    lHeute = Fix(Now)
    cJetzt = Format$(Now, "HH:MM")
    
    txtStatus.Text = ""
    picprogress.Visible = True
    
    iStufe = 2

    cSQL = "Select * from " & sZufall & " order by ARTNR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then 'zufall
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        Do While Not rsrs.EOF
    
            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lcount)
        
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = "-1"
            End If
            
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
            Else
                cLinr = "-1"
            End If
            
            If Not IsNull(rsrs!BESTELLT) Then
                lBestellt = rsrs!BESTELLT
            Else
                lBestellt = 0
            End If
            
            If Not IsNull(rsrs!GELIEFERT) Then
                lBestandZugang = rsrs!GELIEFERT
            Else
                lBestandZugang = 0
            End If
            
            If Not IsNull(rsrs!STCK_PREIS) Then
                dStückPreis = rsrs!STCK_PREIS
            Else
                dStückPreis = 0
            End If
        
            iStufe = 21
            cSQL = "Select * from Artikel where artnr = " & cArtNr
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
            
                If Not IsNull(rsArt!BEZEICH) Then
                    cBezeich = rsArt!BEZEICH
                Else
                    cBezeich = ""
                End If
                
                If Not IsNull(rsArt!KVKPR1) Then
                    dKVkPr1 = rsArt!KVKPR1
                Else
                    dKVkPr1 = 0
                End If
                
                If Not IsNull(rsArt!EAN) Then
                    cEAN = rsArt!EAN
                Else
                    cEAN = ""
                End If

                If Not IsNull(rsArt!BESTAND) Then
                    lBestandAlt = rsArt!BESTAND
                Else
                    lBestandAlt = 0
                End If
                
                If Not IsNull(rsArt!ekpr) Then
                    dStückPreisAlt = rsArt!ekpr
                Else
                    dStückPreisAlt = 0
                End If
                
                If lBestandAlt > 0 Then
                    dGesamtPreisAlt = dStückPreisAlt * lBestandAlt
                Else
                    dGesamtPreisAlt = 0
                End If
                
                dGesamtPreisNeu = dStückPreis * lBestandZugang
                dGesamtPreis = dGesamtPreisAlt + dGesamtPreisNeu
                dGesamtBestand = lBestandAlt + lBestandZugang
                
                dEkpr = 0
                If lBestandAlt < 0 Then
                    dEkpr = dStückPreis
                ElseIf (dGesamtBestand > 0 And lBestandAlt > 0) Or (dGesamtBestand > 0 And lBestandAlt = 0) Then
                    dEkpr = dGesamtPreis / dGesamtBestand
                Else
                    dEkpr = dStückPreis
                End If
                
                rsArt.Edit
                rsArt!lekpr = rsrs!lekpr
                
                If rsArt!GEFUEHRT = "N" Then
                    If gbWEautoGef Then
                        rsArt!GEFUEHRT = "J"
                        insertArtikelDetail lKJADate, cKJAZeit, gcKasNum, CInt(gcBedienerNr), CLng(cArtNr), "gefuehrt", "J"
                    End If
                End If
                
                rsArt!SYNStatus = "E"
                
                If dEkpr <> 0 Then
                    rsArt!ekpr = dEkpr
                End If
                
                rsArt.Update
                
            End If
            rsArt.Close: Set rsArt = Nothing
            
            iStufe = 22
            
            If lBestandZugang <> 0 Then
                Bestandsveraenderung cArtNr, CLng(dGesamtBestand), "WE aus Bestellung"
            End If
            
            If lBestandZugang <> 0 Then
                ABinBESTAKT cArtNr, lBestandZugang, "WE aus Bestellung"
            End If
            
            iStufe = 23
            
            If KundenbestBestätigung(cArtNr, CDbl(lBestandZugang)) = True Then
                Command1(3).Visible = True
                anzeige "ERFOLG", "", Label5
            End If
        
            iStufe = 24
            
            Dim sLieferschein As String
            If Trim(txtLieferschein.Text) <> "" Then
                sLieferschein = Trim(txtLieferschein.Text)
            Else
                sLieferschein = cdatei
            End If
    
            Set rsHis = gdBase.OpenRecordset("ZUGANG", dbOpenTable)
            rsHis.AddNew
            rsHis!artnr = cArtNr
            rsHis!BEZEICH = cBezeich
            rsHis!linr = cLinr
            rsHis!EAN = cEAN
            rsHis!ADATE = lHeute
            rsHis!Uhrzeit = cJetzt
            rsHis!BEDNU = Val(gcBedienerNr)
            rsHis!bedname = gcUserName
            rsHis!FILIALNR = 1
            rsHis!bestandalt = lBestandAlt
            rsHis!BEWEGUNG = lBestandZugang
            rsHis!BESTANDneu = dGesamtBestand
            rsHis!ekpr = dStückPreis
            rsHis!rek = dStückPreis
            rsHis!LS = sLieferschein
            rsHis.Update
            rsHis.Close: Set rsHis = Nothing
             
            Set rsHis = gdBase.OpenRecordset("ZUGANGF", dbOpenTable)
            rsHis.AddNew
            rsHis!artnr = cArtNr
            rsHis!BEZEICH = cBezeich
            rsHis!linr = cLinr
            rsHis!EAN = cEAN
            rsHis!ADATE = lHeute
            rsHis!Uhrzeit = cJetzt
            rsHis!BEDNU = Val(gcBedienerNr)
            rsHis!bedname = gcUserName
            rsHis!FILIALNR = gcFilNr
            rsHis!bestandalt = lBestandAlt
            rsHis!BEWEGUNG = lBestandZugang
            rsHis!BESTANDneu = dGesamtBestand
            rsHis!ekpr = dStückPreis
            rsHis!rek = dStückPreis
            rsHis!LS = sLieferschein
            rsHis!SENDOK = False
            rsHis.Update
            rsHis.Close: Set rsHis = Nothing
            
             
            iStufe = 25
               
            If Not IsNull(rsrs!LIBESNR) Then
                cLiBesNr = rsrs!LIBESNR
            Else
                cLiBesNr = ""
            End If
            
            If Not IsNull(rsrs!lekpr) Then
                IsinArtlief cArtNr, cLinr, rsrs!lekpr, cLiBesNr
            Else
                IsinArtlief cArtNr, cLinr, "0", cLiBesNr
            End If
            
            iStufe = 26
            
            If lBestandZugang > 0 Then
            
                If gbNoETIWeAusBe = False Then
                    If txtLieferschein.Text = "" Then
                        schreibeWKEtidru cArtNr, lBestandZugang, CLng(gcFilNr)
                    Else
                        Set rsEtiLs = gdBase.OpenRecordset("ETIDRULS", dbOpenTable)
                        rsEtiLs.AddNew
                        rsEtiLs!artnr = cArtNr
                        rsEtiLs!BEZEICH = cBezeich
                        rsEtiLs!BESTAND = lBestandZugang
                        rsEtiLs!ANZAHL = lBestandZugang
                        rsEtiLs!vkpr = dKVkPr1
                        rsEtiLs!LIBESNR = cLiBesNr
                        rsEtiLs!EAN = cEAN
                        rsEtiLs!LPZ = 0
                        rsEtiLs!linr = cLinr
                        rsEtiLs!filnr = gcFilNr
                        rsEtiLs!WEDATE = lHeute
                        rsEtiLs!LS = sLieferschein
                        rsEtiLs.Update
                        rsEtiLs.Close: Set rsEtiLs = Nothing
                    End If
                End If
                  
            End If
            
            iStufe = 27
            
            cSQL = "Select * from BESTREST where DATEINAME like '" & cdatei & "*' "
            cSQL = cSQL & " and ARTNR = " & cArtNr & " "
           
            Set rsRest = gdBase.OpenRecordset(cSQL)
            If Not rsRest.EOF Then
                rsRest.MoveFirst

                If lBestellt <= lBestandZugang Then
                    If Not rsRest.EOF Then
                        rsRest.delete
                    End If
                    rsrs.delete
                    
                    cSQL = "Delete from " & cdatei & " where ARTNR = " & cArtNr
                    gdBase.Execute cSQL, dbFailOnError
                Else
                    If Not rsRest.EOF Then
                    
                        rsRest.Edit
                        rsRest!BESTVOR = (lBestellt - lBestandZugang)
                        rsRest!UPD_DATUM = lHeute
                        rsRest.Update
                        
                    End If
                End If
            ElseIf rsRest.EOF And bAnfuegen = True Then
            
                If lBestellt - lBestandZugang > 0 Then
            
                rsRest.AddNew
                rsRest!linr = cAnfuLinr
                rsRest!artnr = cArtNr
                rsRest!lekpr = rsrs!lekpr
                rsRest!BESTVOR = (lBestellt - lBestandZugang)
                rsRest!Dateiname = cdatei & ".DBF"
                rsRest!BEST_DATUM = lHeute
                rsRest!UPD_DATUM = lHeute
                rsRest.Update
                
                End If
            End If
            rsRest.Close
            rsrs.MoveNext
            
            iStufe = 28
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    iStufe = 29
    
    

    'ist es eine Zwischengespeicherte Datei?
    
    Dim bZwischgespeicherte As Boolean
    bZwischgespeicherte = False
    If SpalteInTabellegefundenNEW(cdatei, "BESTELLT", gdBase) Then
        bZwischgespeicherte = True
    End If

    
    
    
    cSQL = "Select * from BESTREST where DATEINAME like '" & cdatei & "*' "
    cSQL = cSQL & " order by ARTNR "
    Set rsRest = gdBase.OpenRecordset(cSQL)
    If Not rsRest.EOF Then
        rsRest.MoveFirst
        Do While Not rsRest.EOF
            If Not IsNull(rsRest!artnr) Then
                cArtNr = rsRest!artnr
            Else
                cArtNr = "-1"
            End If
            
            
            
            
            
            cSQL = "Select * from " & cdatei & " where ARTNR = " & cArtNr
            Set rsBest = gdBase.OpenRecordset(cSQL)
            If Not rsBest.EOF Then
                rsBest.MoveFirst
                rsBest.Edit
                If bZwischgespeicherte = True Then
                    rsBest!BESTELLT = rsRest!BESTVOR
                    rsBest!GELIEFERT = rsRest!BESTVOR
                    rsBest!BERECHNET = rsRest!BESTVOR
                Else
                    rsBest!BESTVOR = rsRest!BESTVOR
                End If
                rsBest.Update
            ElseIf rsBest.EOF And bAnfuegen = True Then
                rsBest.AddNew
                rsBest!artnr = rsRest!artnr
                rsBest!BEZEICH = cAnfuegenBez
                
                
'                rsBest!BESTVOR = rsRest!BESTVOR
                
                If bZwischgespeicherte = True Then
                    rsBest!BESTELLT = rsRest!BESTVOR
                    rsBest!GELIEFERT = rsRest!BESTVOR
                    rsBest!BERECHNET = rsRest!BESTVOR
                Else
                    rsBest!BESTVOR = rsRest!BESTVOR
                End If
                rsBest!linr = rsRest!linr
                rsBest!lekpr = dAnfuLEKPR
                rsBest!LIBESNR = 0
                rsBest.Update
            End If
            rsBest.Close
            rsRest.MoveNext
         Loop
    Else
        cSQL = "Delete from TABDATUM where TABNAME like '" & cdatei & "*' "
        gdBase.Execute cSQL, dbFailOnError
        loeschNEW cdatei, gdBase
    End If
    rsRest.Close
    

    
    If Check7.Value = vbChecked Then
        cSQL = "Update " & cdatei & " inner join  TempZufall" & sZufall
        cSQL = cSQL & " on " & cdatei & ".artnr = TempZufall" & sZufall & ".artnr "
    
    
    
        If bZwischgespeicherte = True Then
    
            cSQL = cSQL & " set " & cdatei & ".GELIEFERT = TempZufall" & sZufall & ".geliefert "
    '        rsBest!BESTELLT = rsRest!BESTVOR
    '        rsBest!GELIEFERT = rsRest!BESTVOR
    '        rsBest!BERECHNET = rsRest!BESTVOR
        Else
            cSQL = cSQL & " set " & cdatei & ".BESTVOR = TempZufall" & sZufall & ".geliefert "
    '        rsBest!BESTVOR = rsRest!BESTVOR
        End If
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    
    
    
    
    iStufe = 30

    If sZufall <> "" Then
        loesch sZufall
    End If
    
    If gbDELBDAT Then
        cSQL = "Delete from BESTREST where DATEINAME like '" & cdatei & "*' "
        gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Delete from TABDATUM where TABNAME = '" & cdatei & "'"
        gdBase.Execute cSQL, dbFailOnError
        
        loeschNEW cdatei, gdBase
    End If
    
    txtStatus.Text = ""
    picprogress.Visible = False
    MsgBox "Die Übernahme der Lieferung wurde erfolgreich abgeschlossen.", vbInformation, "Winkiss Hinweis:"
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EingangDerArtikel"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten." & iStufe

    Fehlermeldung1
End Sub
Private Sub druckendgültigTeil1()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim lRows As Long
    Dim lCols As Long
    Dim cFeld As String
    Dim cFeld1 As String
    Dim cFeld2 As String
    
    
    Screen.MousePointer = 11
    
    
    loeschNEW "DT" & srechnertab, gdBase
    CreateTableT2 "DT" & srechnertab, gdBase
    
    lRows = MSFlexGrid1.Rows
    lCols = MSFlexGrid1.Cols
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld1 = Trim$(Mid(cFeld, 15, 10))
    
    cFeld = UCase$(List2.list(List2.ListIndex))
    cFeld2 = Trim$(Mid(cFeld, 73, 10))
    
    
    
    cSQL = "Insert into DT" & srechnertab & " Select "
    cSQL = cSQL & "'" & gcWEdatei & "'  as BESTELLDAT, "
    cFeld = ""
    cSQL = cSQL & "'" & cFeld2 & "'  as BESTELLNR, "
    cFeld = ""
    cSQL = cSQL & "'" & cFeld1 & "'  as BESTELLUNG, "
    cFeld = Label3(7).Caption
    cSQL = cSQL & "'" & cFeld & "' as LIEFERANT, "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , LEKPR "
    cSQL = cSQL & " , BESTELLT "
    cSQL = cSQL & " , GELIEFERT "
    cSQL = cSQL & " , BERECHNET "
    cSQL = cSQL & " , LIEFBETRAG "
    cSQL = cSQL & " , Zeilen_rab as ZEILENRAB "
    cSQL = cSQL & " , ZEILENWERT "
    cSQL = cSQL & " , Rechn_rab as RECHNRAB "
    cSQL = cSQL & " , Rechn_wert as RECHNWERT "
    cSQL = cSQL & " , Stck_Preis as STCKPREIS "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , 0 as LAGERP "
    cSQL = cSQL & " from " & sZufall & " "
    cSQL = cSQL & cSort
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update DT" & srechnertab & " inner join LAGERPLATZ on val(DT" & srechnertab & ".ARTNR) = LAGERPLATZ.ARTNR "
    cSQL = cSQL & " set DT" & srechnertab & ".LAGERP = LAGERPLATZ.LAGERP "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update DT" & srechnertab & " set LAGERP = 0 where lagerp is null "
    gdBase.Execute cSQL, dbFailOnError

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "druckendgültigTeil1"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung  ist ein Fehler aufgetreten."

    Fehlermeldung1
    
   
End Sub
Private Sub druckendgültigTeil2()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    Dim cSQL As String
    Dim cPfad1 As String
    Dim rsrs As DAO.Recordset
    Dim cNRRegel As String
    Dim cKondiLief As String
    Dim cKondiArtnr As String
    
    Dim lKonditionLi As Long
    Dim lFaktorLi As Long
    Dim lKonditionAr As Long
    Dim lFaktorAr As Long
    
    Dim lBestellt As Long
    Dim lGeliefert As Long
    Dim lBerechnet As Long
    Dim lNaturalAnz As Long
    Dim lGeliefertPlusNr As Long
    Dim lFehlerfarbe As Long
    
    Dim cArtNr As String
    Dim lMulti As Long
    
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    loeschNEW "DRU_TEMP", gdApp
    CreateTableT2 "DRU_TEMP", gdApp
    
    SpalteAnfuegenNEW "DRU_TEMP", "LAGERP", "Double", gdApp
    
    loeschNEW "DT" & srechnertab, gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "DT" & srechnertab
    
    cSQL = "Insert into DRU_TEMP select * from DT" & srechnertab
    gdApp.Execute cSQL, dbFailOnError
    
    SpalteAnfuegenNEW "DRU_TEMP", "HFarbe", "LONG", gdApp
    SpalteAnfuegenNEW "DRU_TEMP", "NRREGEL", "Text(5)", gdApp
   
    
    cKondiLief = ""
'    cKondiLief = HolekondiLief(Label3(6).Caption)
'    If cKondiLief <> "" Then
'        Dim sArray() As String
'        sArray = Split(cKondiLief, ";")
'        lKonditionLi = sArray(0)
'        lFaktorLi = sArray(1)
'    End If
    
    
    cSQL = "Select * from DRU_TEMP "
    Set rsrs = gdApp.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            lBestellt = 0
            lGeliefert = 0
            lBerechnet = 0
            lNaturalAnz = 0
            lGeliefertPlusNr = 0
            cNRRegel = ""
            cArtNr = ""
            lKonditionAr = 0
            lFaktorAr = 0
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            cKondiArtnr = HolekondiArtnr(cArtNr)
            If cKondiArtnr = "" Then cKondiArtnr = cKondiLief
            
            If cKondiArtnr <> "" Then
                cNRRegel = SwapStr(cKondiArtnr, ";", "+")
                
                Dim sArrayArt() As String
                sArrayArt = Split(cKondiArtnr, ";")
                lKonditionAr = sArrayArt(0)
                lFaktorAr = sArrayArt(1)
            Else
                lKonditionAr = 0
                lFaktorAr = 0

            End If
            
            If Not IsNull(rsrs!BESTELLT) Then
                lBestellt = rsrs!BESTELLT
            End If
            
            If Not IsNull(rsrs!GELIEFERT) Then
                lGeliefert = rsrs!GELIEFERT
                lGeliefertPlusNr = lGeliefert
            End If
            
            If Not IsNull(rsrs!BERECHNET) Then
                lBerechnet = rsrs!BERECHNET
            End If
            
            'Beispiel 6 bestellt und 6 geliefert NR Regel = 6+1
            'Erstens gibt es eine Regel
            
            If lKonditionAr > 0 Then
                'Zweitens Wurde überhaupt etwas geliefert?
                If lGeliefert > 0 Then
                    'Drittens ist geliefert >= bestellt?
                    If lGeliefert >= lBestellt Then
                    
                        '5. führt die Bestellhöhe überhaupt zum Naturalrabat
                        If lBestellt >= lKonditionAr Then
                        
                            '6. Na dann rechne mal den Naturalrabatt bei der Bestellmenge aus
                            lMulti = lBestellt / lKonditionAr
                            lMulti = Fix(lMulti)
                            
                            lNaturalAnz = lMulti * lFaktorAr
                            
                            lGeliefertPlusNr = lBestellt + lNaturalAnz
                            
                        End If
                    
                    End If
                End If
            End If
            
            If lGeliefert >= lGeliefertPlusNr Then
                'alles Ok
                '= weiß
                lFehlerfarbe = vbWhite
            Else
                'entgangener NR
                '= rot
                lFehlerfarbe = vbRed
            End If
            
            If lBestellt = 0 And lGeliefert > 0 Then
                'Lila
                lFehlerfarbe = &HFF00FF
            End If
            
            If lGeliefert < lBestellt Then
                lFehlerfarbe = vbRed
            End If
            
            If lBestellt > 0 And lGeliefert = 0 Then
                'gelb
                lFehlerfarbe = vbYellow
            End If
            
            
            
            rsrs.Edit
            rsrs!NRREGEL = cNRRegel
            rsrs!HFarbe = lFehlerfarbe
            rsrs.Update
            
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Pause (1)
    reportbildschirmApp "", "aWKL15ae"

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "druckendgültigTeil2"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung  ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function HolekondiLief(cLinr As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    
    HolekondiLief = ""

    cSQL = "Select * from KONDITIONENL where LINR = " & cLinr & " "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!kondi) Then
            HolekondiLief = rsrs!kondi & ";"
        End If

        If Not IsNull(rsrs!Faktor) Then
            HolekondiLief = HolekondiLief & rsrs!Faktor
        End If
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HolekondiLief"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function isArtnrEnthalten(cArtNr As String, cdatei As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    
    isArtnrEnthalten = False

    cSQL = "Select * from " & cdatei & " where ARTNR = " & cArtNr & " "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        isArtnrEnthalten = True
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "isArtnrEnthalten"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function HolekondiArtnr(cArtNr As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    
    HolekondiArtnr = ""

    cSQL = "Select * from KONDITIONEN where ARTNR = " & cArtNr & " "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!kondi) Then
            HolekondiArtnr = rsrs!kondi & ";"
        End If

        If Not IsNull(rsrs!Faktor) Then
            HolekondiArtnr = HolekondiArtnr & rsrs!Faktor
        End If
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "HolekondiArtnr"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Check4_Click()
    On Error GoTo LOKAL_ERROR
    
    If Check4.Value = vbChecked Then
        cboStrichEndlos.Visible = True
        cboStrichEndlos.Refresh
        setzedrucker gcEtikettenDrucker
        
    Else
        cboStrichEndlos.Visible = False
        cboStrichEndlos.Refresh
        setzedrucker gcListenDrucker
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check4_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text5.BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text7_LostFocus()
On Error GoTo LOKAL_ERROR
    
    Text7.BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text7_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


