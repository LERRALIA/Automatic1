VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWKL24 
   BackColor       =   &H00C0C000&
   Caption         =   "Kreditverwaltung"
   ClientHeight    =   8910
   ClientLeft      =   1455
   ClientTop       =   1875
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
   Icon            =   "frmWKL24.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8910
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
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
      Height          =   1215
      Left            =   10560
      TabIndex        =   69
      Top             =   -120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text21 
         Height          =   255
         Index           =   2
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   7320
         Width           =   1335
      End
      Begin VB.TextBox Text21 
         Height          =   255
         Index           =   1
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   6960
         Width           =   1335
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   6
         Left            =   7560
         TabIndex        =   87
         Top             =   6960
         Width           =   1815
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Frame Frame6 
         Caption         =   "anzeigen"
         Height          =   975
         Left            =   2880
         TabIndex        =   83
         Top             =   120
         Width           =   3255
         Begin VB.OptionButton Option1 
            Caption         =   "bezahlte"
            Height          =   210
            Index           =   7
            Left            =   240
            TabIndex        =   86
            Top             =   720
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            Caption         =   "nicht bezahlte"
            Height          =   210
            Index           =   5
            Left            =   240
            TabIndex        =   85
            Top             =   480
            Width           =   2775
         End
         Begin VB.OptionButton Option1 
            Caption         =   "alle"
            Height          =   210
            Index           =   4
            Left            =   240
            TabIndex        =   84
            Top             =   240
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "sortiert nach"
         Height          =   975
         Left            =   6360
         TabIndex        =   79
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton Option1 
            Caption         =   "Zahlungsziel"
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   82
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Kundennummer"
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   81
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rechnungsnummer"
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   80
            Top             =   480
            Width           =   2175
         End
      End
      Begin sevCommand3.Command Command3 
         Height          =   375
         Index           =   12
         Left            =   720
         TabIndex        =   76
         Top             =   720
         Width           =   1575
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
         Caption         =   "aktualisieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   75
         Text            =   "14"
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Eigene Anschrift im Rechnungsfuß"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   6480
         Visible         =   0   'False
         Width           =   3495
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Index           =   10
         Left            =   9480
         TabIndex        =   70
         Top             =   6960
         Width           =   1815
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
         Height          =   5295
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   9
         FixedCols       =   2
         BackColorSel    =   16711680
         ForeColorSel    =   65535
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin sevCommand3.Command Command98 
         Height          =   360
         Left            =   10800
         TabIndex        =   93
         Top             =   120
         Width           =   405
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
         ToolTip         =   "Spaltenanordung der Tabelle bestimmen"
         ToolTipTitle    =   "Spaltenanordung"
         ButtonStyle     =   2
         Caption         =   ""
         Filename        =   "D:\Thomas\VB6\Winkiss\Zubehör\tab24.gif"
         Picture         =   "frmWKL24.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   2
         Left            =   6840
         TabIndex        =   94
         ToolTipText     =   "Kalender"
         Top             =   6960
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   1
         Left            =   6840
         TabIndex        =   95
         ToolTipText     =   "Kalender"
         Top             =   7320
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
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Label label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   4920
         TabIndex        =   92
         Top             =   7320
         Width           =   375
      End
      Begin VB.Label label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFF00&
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   4800
         TabIndex        =   91
         Top             =   6960
         Width           =   495
      End
      Begin VB.Label label1 
         Alignment       =   1  'Rechts
         Height          =   255
         Index           =   8
         Left            =   6000
         TabIndex        =   88
         Top             =   6600
         Width           =   5295
      End
      Begin VB.Label lbl8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   4560
         TabIndex        =   78
         Top             =   6480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Zahlungsziel in Tagen:"
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "-1"
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
         Left            =   6000
         TabIndex        =   74
         Top             =   7800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblUeber 
         BackStyle       =   0  'Transparent
         Caption         =   "Offene Postenliste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   73
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   5640
      TabIndex        =   62
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picprogress 
      Height          =   375
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   4755
      TabIndex        =   61
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Frame Frame3 
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
      Height          =   975
      Left            =   11160
      TabIndex        =   27
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   5
         Left            =   2520
         TabIndex        =   64
         Top             =   6840
         Width           =   2055
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
         Caption         =   "Mahnung drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   4
         Left            =   6960
         TabIndex        =   36
         Top             =   6840
         Width           =   1215
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
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   3
         Left            =   8280
         TabIndex        =   31
         Top             =   6840
         Width           =   1455
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
         Caption         =   "Rechnung löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   2
         Left            =   4680
         TabIndex        =   30
         Top             =   6840
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
         Caption         =   "alte Rechnung rückgängig machen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   1
         Left            =   9840
         TabIndex        =   32
         Top             =   6840
         Width           =   1455
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
      Begin sevCommand3.Command Command3 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   6840
         Width           =   2295
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
         Caption         =   "alte Rechnungen ansehen / drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
         Height          =   6135
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   10821
         _Version        =   393216
         Cols            =   9
         FixedCols       =   2
         BackColorSel    =   16711680
         ForeColorSel    =   65535
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Eigene Anschrift im Rechnungsfuß"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   6480
         Width           =   3495
      End
      Begin VB.Label lblUeber 
         BackStyle       =   0  'Transparent
         Caption         =   "vorhandene Rechnungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label Label4 
         Caption         =   "-1"
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
         Left            =   6000
         TabIndex        =   33
         Top             =   7800
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
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
      Height          =   8055
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   11400
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   7
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   65
         Top             =   5760
         Width           =   495
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   2
         Left            =   9660
         TabIndex        =   23
         Top             =   6960
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   661
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
            Size            =   12
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
         Index           =   3
         Left            =   4770
         TabIndex        =   22
         Top             =   6960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
            Size            =   12
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
         Caption         =   "Zusatz Rechnung 1"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   1
         Left            =   1960
         TabIndex        =   21
         Top             =   6960
         Width           =   2780
         _ExtentX        =   4895
         _ExtentY        =   661
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
            Size            =   12
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
         Caption         =   "Rechnung schreiben"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   6960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
            Size            =   12
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
         Caption         =   "Kunde bezahlt"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         Caption         =   "alle markieren"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   7800
         MaxLength       =   35
         TabIndex        =   54
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   9720
         MaxLength       =   35
         TabIndex        =   50
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   840
         MaxLength       =   35
         TabIndex        =   51
         Top             =   6120
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   3
         Left            =   5160
         MaxLength       =   8
         TabIndex        =   52
         Top             =   6120
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   4
         Left            =   6720
         MaxLength       =   35
         TabIndex        =   53
         Top             =   6120
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   5
         Left            =   2640
         MaxLength       =   35
         TabIndex        =   48
         Top             =   5760
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   6
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   49
         Top             =   5760
         Width           =   735
      End
      Begin MSComDlg.CommonDialog cdlprinter 
         Left            =   1800
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4455
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   9
         BackColorSel    =   16711680
         ForeColorSel    =   65535
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox Check3 
         Caption         =   "ohne Druck"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   6480
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Eigene Anschrift im Rechnungsfuß"
         Height          =   255
         Left            =   2280
         TabIndex        =   66
         Top             =   6480
         Width           =   3495
      End
      Begin sevCommand3.Command Command2 
         Height          =   375
         Index           =   4
         Left            =   7210
         TabIndex        =   96
         Top             =   6960
         Width           =   2425
         _ExtentX        =   4286
         _ExtentY        =   661
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
            Size            =   12
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
         Caption         =   "Zusatz Rechnung 2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "-1"
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
         Left            =   8040
         TabIndex        =   101
         Top             =   7680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-1"
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
         Left            =   7440
         TabIndex        =   100
         Top             =   8160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-1"
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
         Left            =   7440
         TabIndex        =   99
         Top             =   7920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-1"
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
         Left            =   7440
         TabIndex        =   98
         Top             =   7680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-1"
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
         Left            =   6840
         TabIndex        =   97
         Top             =   7680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblUeber 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   4
         Left            =   4560
         TabIndex        =   58
         Top             =   0
         Width           =   6855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   57
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname"
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
         Index           =   7
         Left            =   2640
         TabIndex        =   56
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname"
         Height          =   255
         Index           =   3
         Left            =   6960
         TabIndex        =   55
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Firma"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   47
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anrede"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   46
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Index           =   2
         Left            =   9120
         TabIndex        =   45
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Straße"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Plz"
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   43
         Top             =   6120
         Width           =   255
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Ort"
         Height          =   255
         Index           =   6
         Left            =   6000
         TabIndex        =   42
         Top             =   6120
         Width           =   615
      End
      Begin VB.Label lblUeber 
         BackStyle       =   0  'Transparent
         Caption         =   "Rechnungsanschrift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label lblUeber 
         BackStyle       =   0  'Transparent
         Caption         =   "Kunde mit offenen Krediten (Detail)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   0
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   "-1"
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
         Left            =   6120
         TabIndex        =   35
         Top             =   8160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "-1"
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
         Left            =   6120
         TabIndex        =   34
         Top             =   7920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "abzurechnende Position(en) durch Anklicken der Zeile auf Status ""ausbuchen"" stellen ! Erneutes Anklicken setzt Status zurück."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   10935
      End
      Begin VB.Label Label3 
         Caption         =   "-1"
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
         Left            =   6120
         TabIndex        =   25
         Top             =   7680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   10440
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   7440
         TabIndex        =   18
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   17
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "offen"
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
         Left            =   10440
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ort"
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
         Left            =   7440
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "PLZ"
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
         Left            =   6240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Straße"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "KundNr"
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
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
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
      Height          =   2055
      Left            =   6240
      TabIndex        =   1
      Top             =   3120
      Width           =   2535
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   4
         Left            =   5370
         TabIndex        =   68
         Top             =   6960
         Width           =   2415
         _ExtentX        =   4260
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
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "offene Posten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "alle anzeigen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9360
         TabIndex        =   60
         Top             =   0
         Width           =   1935
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   3
         Left            =   7820
         TabIndex        =   4
         Top             =   6960
         Width           =   1750
         _ExtentX        =   3096
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
            Name            =   "Arial"
            Size            =   12
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
         Height          =   495
         Index           =   2
         Left            =   2920
         TabIndex        =   3
         Top             =   6960
         Width           =   2415
         _ExtentX        =   4260
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
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "alte Rechnungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   5
         Top             =   6960
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
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
            Name            =   "Arial"
            Size            =   12
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
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   6960
         Width           =   2765
         _ExtentX        =   4868
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
            Name            =   "Arial"
            Size            =   12
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
         Caption         =   "nächster Schritt"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6495
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   11456
         _Version        =   393216
         Cols            =   9
         FixedCols       =   2
         RowHeightMin    =   200
         BackColor       =   16777215
         BackColorSel    =   16711680
         ForeColorSel    =   65535
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblUeber 
         BackStyle       =   0  'Transparent
         Caption         =   "Kunden mit offenen Krediten der letzten 61 Tage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label0 
         Caption         =   "-1"
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
         Left            =   7680
         TabIndex        =   6
         Top             =   7560
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   11520
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kreditverwaltung"
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
      Left            =   360
      TabIndex        =   37
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmWKL24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bUnterschiedlicheRechnunganschriften    As Boolean
Dim cFirma                                  As String
Dim cAnrede                                 As String
Dim cTitel                                  As String
Dim SpaltennummerZahlZiel                   As Byte
Dim SpaltennummerReDatum                    As Byte
Dim SpaltennummerKdNr                       As Byte
Dim SpaltennummerReNr                       As Byte
Dim SpaltennummerStatusBez                  As Byte
Dim SpaltennummerBezahlInfo                 As Byte
Dim gbAenderBezahlInfo                      As Boolean
Private Sub Check2_Click()
On Error GoTo LOKAL_ERROR
    
    If Check2.value = vbChecked Then
        LeseOffeneKrediteWKL24 0
    Else
        LeseOffeneKrediteWKL24 61
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command0_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case index
    
        Case Is = 2
            Text21(1).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            Text21(2).Text = Text21(1).Text
            
            Text21(2).SetFocus
        Case Is = 1
            Text21(2).Text = Format(Datumschreiben11a(3500, 340), "DD.MM.YY")
            'fertig
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command98_Click()
On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "RENR"
    gstab = "OFPO"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command98_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "DRU_ALKR", gdBase
    loeschNEW "DRU_REKO", gdBase
    loeschNEW "DRU_REPO", gdBase
    loeschNEW "RE_EXPORT", gdBase
    
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
Private Sub PositionierenWKL24()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 840
    Frame1.Left = 240
    Frame1.Height = 8055
    Frame1.Width = 11415
    
    Frame2.Top = 840
    Frame2.Left = 240
    Frame2.Height = 8055
    Frame2.Width = 11415
    
    Frame3.Top = 840
    Frame3.Left = 240
    Frame3.Height = 8055
    Frame3.Width = 11415
    
    Frame4.Top = 840
    Frame4.Left = 240
    Frame4.Height = 8055
    Frame4.Width = 11415
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub AusbuchenKreditVerkaufWKL24()
    On Error GoTo LOKAL_ERROR
    
    Dim cFlag As String
    Dim cKundnr As String
    Dim cDatum As String
    Dim cArtNr As String
    Dim cBezeich As String
    Dim cAnzahl As String
    Dim dEPreis As Double
    Dim dGPreis As Double
    
    Dim lDatum As Long
    Dim cSQL As String
    Dim iRet As Integer
    Dim dZugang As Double
    Dim dAbgang As Double
    Dim ctmp As String
    Dim cSTATUS As String
    Dim rsrs As Recordset
    Dim lDatumJetzt As Long
    Dim cZeitJetzt As String
    
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim lRet As Long
    
    'Kassenbon für Kredittilgung produzieren
    Dim cDaten As String
    Dim aDeviceName As String
    Dim lAnzZeile As Long
    ReDim cDruckZeile(1 To 1) As String
    Dim iLenZeile As Integer
    Dim lcount As Long
    Dim dSumme As Double
    Dim cDrucker As String
    Dim bReturn As Boolean
    Dim lAnz As Long
    
    Dim lKJADate As Long
    Dim cKJAZeit As String
    
    
    'Variablen für die Schublade
    
    Dim dAktZeit            As Double
    Dim dNeuZeit            As Double
    Dim cEscapeSequenz      As String
    
    cKundnr = Label2(0).Caption
    
    lAnzRecords = MSFlexGrid2.Rows
    
    cSQL = "Wollen Sie den Kreditkauf des " & vbCrLf & vbCrLf
    cSQL = cSQL & "Kunden " & UCase$(Label2(1).Caption) & " ( = " & cKundnr & ") " & vbCrLf & vbCrLf
    cSQL = cSQL & "wirklich ausbuchen? "
    
    iRet = MsgBox(cSQL, vbQuestion + vbYesNo, "AUSBUCHEN")
    If iRet = vbYes Then
        frmWKL28.Show 1
        
        If gcKreditKarte = "" Then
            Exit Sub
        End If
        
    Else
        Exit Sub
    End If
    
    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz
    
    'Schublade öffnen
SCHUBLADE:

    'Bon-Drucker setzen
    iLenZeile = 32
    aDeviceName = gcBonDrucker
    
    setzedrucker gcBonDrucker
    


    '*** Kopfdaten setzen *********************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(0)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cDaten = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To 1) As String
        cDruckZeile(lAnzZeile) = cDaten
    End If
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(1)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cDaten = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cDaten
    End If
    
    '******************************************************************
            
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO-VERSION!"
    Else
        cDaten = gcBonText(4)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cDaten = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cDaten
    End If
    
    '***********************************************
    'Kopfdaten 4.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION"
    Else
        cDaten = gcBonText(12)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '******************************************************************
            
    cDaten = String$(iLenZeile, "-")
    KonvertAnsiAscii cDaten
    cDaten = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cDaten
    
    '******************************************************************
    
    cDaten = "  K R E D I T - T I L G U N G"
    KonvertAnsiAscii cDaten
    cDaten = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cDaten
    
    '******************************************************************
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = " " & vbCrLf
        
    '******************************************************************
            
            
            
    For lAktRecord = 1 To lAnzRecords - 1
    
        MSFlexGrid2.Row = lAktRecord
        MSFlexGrid2.Col = 0
        cSTATUS = MSFlexGrid2.Text
        
        If cSTATUS = "ausbuchen" Then
            MSFlexGrid2.Col = 1
            cDatum = MSFlexGrid2.Text
            lDatum = DateValue(cDatum)
            
            MSFlexGrid2.Col = 2
            cArtNr = MSFlexGrid2.Text
            
            MSFlexGrid2.Col = 3
            cBezeich = MSFlexGrid2.Text
            
            MSFlexGrid2.Col = 4
            cAnzahl = MSFlexGrid2.Text
            
            MSFlexGrid2.Col = 5
            ctmp = MSFlexGrid2.Text
            ctmp = fnMoveComma2Point$(ctmp)
            dEPreis = Val(ctmp)
            
            MSFlexGrid2.Col = 6
            ctmp = MSFlexGrid2.Text
            ctmp = fnMoveComma2Point$(ctmp)
            dGPreis = Val(ctmp)
            
            MSFlexGrid2.Col = 9
            cFlag = MSFlexGrid2.Text
           
            
            '******************************************************************
            
            cDaten = "Kauf-Datum:           " & cDatum
            KonvertAnsiAscii cDaten
            cDaten = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cDaten
            
            '******************************************************************
            
            ctmp = cArtNr
            ctmp = Trim$(ctmp)
            ctmp = ctmp & Space$(6 - Len(ctmp))
            cDaten = ctmp & " "
            
            ctmp = cBezeich
            If Len(ctmp) > 25 Then
                ctmp = Left(ctmp, 25)
            End If
            cDaten = cDaten & ctmp
            
            KonvertAnsiAscii cDaten
            cDaten = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cDaten
            
            '******************************************************************
            
            
            ctmp = cAnzahl
            ctmp = Trim$(ctmp)
            ctmp = ctmp & Space$(5 - Len(ctmp))
            cDaten = ctmp & " à "
            
            
            ctmp = Format$(dEPreis, "#####0.00")
            ctmp = Space$(11 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp & "  "
            
            ctmp = Format$(dGPreis, "#####0.00")
            ctmp = Space$(11 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            
            dSumme = dSumme + dGPreis
            
            KonvertAnsiAscii cDaten
            cDaten = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cDaten
            
            '******************************************************************
            
            cDaten = " " & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cDaten
            
            '******************************************************************
            lDatumJetzt = Fix(Now)
            cZeitJetzt = Format$(Now, "HH:MM:SS")
                
            'Schreibe Tagesprotokoll in AFCBUCH (giZahlArt und gcKreditKarte)
            
            cSQL = "Select * from AFCSTAT where KASNUM = " & gcKasNum
            cSQL = cSQL & " and ADATE = " & lDatumJetzt
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.Edit
            Else
                rsrs.AddNew
            End If
            
            Select Case giZahlArt
                Case Is = 6     'Scheck
                    If Not IsNull(rsrs!TILGSCH) Then
                        rsrs!TILGSCH = rsrs!TILGSCH + dGPreis
                    Else
                        rsrs!TILGSCH = dGPreis
                    End If
                Case Is = 8     'Bar
                    If Not IsNull(rsrs!TILGBAR) Then
                        rsrs!TILGBAR = rsrs!TILGBAR + dGPreis
                    Else
                        rsrs!TILGBAR = dGPreis
                    End If
                Case Is = 17    'Kreditkarte
                    If Not IsNull(rsrs!TILGKAR) Then
                        rsrs!TILGKAR = rsrs!TILGKAR + dGPreis
                        schreibeProtokollUNITXT CStr(dGPreis), "Kartenzahlung"
                    Else
                        rsrs!TILGKAR = dGPreis
                        schreibeProtokollUNITXT CStr(dGPreis), "Kartenzahlung"
                    End If
            End Select
            rsrs!KASNUM = gcKasNum
            rsrs!ADATE = lDatumJetzt
            
            
            If Not IsNull(rsrs!BELEGNR) Then
                If gdBonNr < CLng(rsrs!BELEGNR) Then
                    
                Else
                    rsrs!BELEGNR = gdBonNr
                End If
            Else
                rsrs!BELEGNR = gdBonNr
            End If
            
            rsrs.Update
            rsrs.Close: Set rsrs = Nothing
            
            
            Dim RsKred As DAO.Recordset
            cSQL = "Select top 1 Artnr from KREDIT where "
            cSQL = cSQL & "ARTNR = " & cArtNr & " and "
            cSQL = cSQL & "MENGE = " & cAnzahl & " and "
            cSQL = cSQL & "KUNDNR = " & cKundnr & " and "
            cSQL = cSQL & "ADATE = " & Trim$(Str$(lDatum)) & " "
            
            Set RsKred = gdBase.OpenRecordset(cSQL)
            
            If Not RsKred.EOF Then
                RsKred.delete
            End If
            RsKred.Close: Set RsKred = Nothing
            
            'In die KreditProtokoll für die Zentrale
            
            lKJADate = Fix(Now)
            cKJAZeit = Format$(Now, "HH:MM:SS")
            
            insertKreditZA lKJADate, cKJAZeit, lDatum, CInt(gcBedienerNr), gcKreditKarte, CLng(cKundnr), CLng(cArtNr), CInt(cAnzahl), CStr(gdBonNr)
            
            'Kreditprotokoll ende
            
            MSFlexGrid2.Col = 0
            MSFlexGrid2.Text = "gelöscht"
            
            MSFlexGrid2.Col = 10
            MSFlexGrid2.Text = ""
            
            MSFlexGrid2.Col = 0
            
        End If
    Next lAktRecord
    
    
    '*** Zahlungsart setzen *********************************************
    
    Select Case giZahlArt
        Case Is = 6
            cDaten = "Tilgung per Scheck"
        Case Is = 8
            cDaten = "Tilgung per Barzahlung"
        Case Is = 17
            cDaten = "Tilgung per Kreditkarte (" & gcKreditKarte & ")"
    End Select
    
    KonvertAnsiAscii cDaten
    cDaten = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cDaten

    '*** Fußdaten setzen *********************************************
    
    cDaten = String$(iLenZeile, gsSTERNZEICH)
'    cDaten = String$(iLenZeile, "*")
    KonvertAnsiAscii cDaten
    cDaten = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cDaten



    '*** Tilgungssumme in Euro *********************************************
    
    ctmp = Format$((dSumme), "#####0.00")
    ctmp = Space$(10 - Len(ctmp)) & ctmp
    
    cDaten = "Tilgungsbetrag Euro:  " & ctmp
    KonvertAnsiAscii cDaten
    cDaten = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cDaten



    '*** Fußdaten setzen *********************************************
    
    cDaten = String$(iLenZeile, gsSTERNZEICH)
'    cDaten = String$(iLenZeile, "*")
    KonvertAnsiAscii cDaten
    cDaten = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cDaten
    
    'Hier bauen wir Kunde ein
  
    ctmp = "Ihre KundenNr: " & Label2(0).Caption
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    If gbKUNDENA = True Then
        
        If gbKUIBONfirma Then
            ctmp = lookingForKundendaten(Trim(Label2(0).Caption)).firma
        
            If ctmp <> "" Then
                If Len(ctmp) > 32 Then ctmp = Left(ctmp, 32)
                cDaten = ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        End If
            
        If gbKUIBONtitel Then
            ctmp = lookingForKundendaten(Trim(Label2(0).Caption)).titel
        
            If ctmp <> "" Then
                If Len(ctmp) > 32 Then ctmp = Left(ctmp, 32)
                cDaten = ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        End If
            
        ctmp = ""
        If gbKUIBONvorname Then
            ctmp = lookingForKundendaten(Trim(Label2(0).Caption)).vorname
        End If
        
        If gbKUIBONname Then
            If ctmp = "" Then
                ctmp = ctmp & ""
            Else
                ctmp = ctmp & " "
            End If
            ctmp = ctmp & lookingForKundendaten(Trim(Label2(0).Caption)).nachname
        End If
    
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        If gbKUIBONstrasse Then
            ctmp = lookingForKundendaten(Trim(Label2(0).Caption)).strasse
        
            If Len(ctmp) > 32 Then
                ctmp = Left(ctmp, 32)
            End If
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
            
        ctmp = ""
        If gbKUIBONplz Then
            ctmp = lookingForKundendaten(Trim(Label2(0).Caption)).Plz
        End If
        
        If gbKUIBONort Then
            If ctmp = "" Then
                ctmp = ctmp & ""
            Else
                ctmp = ctmp & " "
            End If
            ctmp = ctmp & lookingForKundendaten(Trim(Label2(0).Caption)).Ort
        End If
        
        If Len(ctmp) > 32 Then
            ctmp = Left(ctmp, 32)
        End If
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
                
        If gbKUIBONtel Then
            ctmp = lookingForKundendaten(Trim(Label2(0).Caption)).telefon
            If ctmp <> "" Then
                cDaten = "Tel " & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        End If
            
        If gbKUIBONmobil Then
            ctmp = lookingForKundendaten(Trim(Label2(0).Caption)).Mobiltel
            If ctmp <> "" Then
                cDaten = "Mobil " & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        End If
        
    End If
    'Kunde ende
    
    '***********************************************
    'Zeile Datum, BelegNr, Uhrzeit drucken
    '***********************************************
    
    ctmp = Format$(Date, "DD.MM.YYYY")
    cDaten = ctmp
    ctmp = Format$(Now, "HH:MM")
    cDaten = cDaten & Space$(4) & ctmp
    
    ctmp = Format$(gdBonNr, "#####0")
    If gbSPIEGEL Then
    
        Dim ctmp111 As String
        Dim N As Integer
        ctmp111 = ctmp
        ctmp = ""
        For N = Len(ctmp111) To 1 Step -1
        
            ctmp = ctmp & Mid(ctmp111, N, 1)
        
        Next N
    End If
    
    ctmp = gcKasNum & "/" & ctmp
    ctmp = Space$(8 - Len(ctmp)) & ctmp
    cDaten = cDaten & Space$(4) & ctmp
    
    
    KonvertAnsiAscii cDaten
    
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(2)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cDaten = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cDaten
    End If
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
    '    cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(3)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cDaten = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cDaten
    End If
    '******************************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
    '    cDaten = "Carsten Schröder"
    Else
        cDaten = gcBonText(5)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cDaten = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cDaten
    End If
    '******************************************************************
    
    For lcount = 1 To 9
        cDaten = " " & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cDaten
        
    Next lcount
    
    '******************************************************************
    For lcount = 1 To 2
        If gbAPI Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
            'Kassenbon abschneiden
            cDaten = Chr$(27) + Chr$(105)
            OpenDrawer aDeviceName, cDaten
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
        
        If lcount = 1 Then
            'Bon-Daten sichern
            SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False
        End If
    Next lcount
    
    
    giZahlArt = 0
    gcKreditKarte = ""
    
    Erase cDruckZeile
    
    'Schublade
    If gbLadeCom Then
        OpenDrawerViaComPortModul20
    Else
        If gbAPI = True Then
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcLade
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AusbuchenKreditVerkaufWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub AusbuchenKreditVerkaufWKL24ohneBon()
    On Error GoTo LOKAL_ERROR
    
    Dim cFlag As String
    Dim cKundnr As String
    Dim cDatum As String
    Dim cArtNr As String
    Dim cBezeich As String
    Dim cAnzahl As String
    Dim dEPreis As Double
    Dim dGPreis As Double
    
    Dim lDatum As Long
    Dim cSQL As String
    Dim iRet As Integer
    Dim dZugang As Double
    Dim dAbgang As Double
    Dim ctmp As String
    Dim cSTATUS As String
    Dim rsrs As Recordset
    Dim lDatumJetzt As Long
    Dim cZeitJetzt As String
    
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim lRet As Long
    
    Dim lKJADate As Long
    Dim cKJAZeit As String
    
    cKundnr = Label2(0).Caption
    
    lAnzRecords = MSFlexGrid2.Rows
    
    cSQL = "Wollen Sie den Kreditkauf des " & vbCrLf & vbCrLf
    cSQL = cSQL & "Kunden " & UCase$(Label2(1).Caption) & " ( = " & cKundnr & ") " & vbCrLf & vbCrLf
    cSQL = cSQL & "wirklich ausbuchen? "
    
    iRet = MsgBox(cSQL, vbQuestion + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        frmWKL28.Show 1
    Else
        Exit Sub
    End If
    
    For lAktRecord = 1 To lAnzRecords - 1
    
        MSFlexGrid2.Row = lAktRecord
        MSFlexGrid2.Col = 0
        cSTATUS = MSFlexGrid2.Text
        
        If cSTATUS = "ausbuchen" Then
            MSFlexGrid2.Col = 1
            cDatum = MSFlexGrid2.Text
            lDatum = DateValue(cDatum)
            
            MSFlexGrid2.Col = 2
            cArtNr = MSFlexGrid2.Text
            
            MSFlexGrid2.Col = 3
            cBezeich = MSFlexGrid2.Text
            
            MSFlexGrid2.Col = 4
            cAnzahl = MSFlexGrid2.Text
            
            MSFlexGrid2.Col = 5
            ctmp = MSFlexGrid2.Text
            ctmp = fnMoveComma2Point$(ctmp)
            dEPreis = Val(ctmp)
            
            MSFlexGrid2.Col = 6
            ctmp = MSFlexGrid2.Text
            ctmp = fnMoveComma2Point$(ctmp)
            dGPreis = Val(ctmp)
            
            MSFlexGrid2.Col = 9
            cFlag = MSFlexGrid2.Text
           
            
           
            
            
            lDatumJetzt = Fix(Now)
            cZeitJetzt = Format$(Now, "HH:MM:SS")
                
            'Schreibe Tagesprotokoll in AFCBUCH (giZahlArt und gcKreditKarte)
            
            cSQL = "Select * from AFCSTAT where KASNUM = " & gcKasNum
            cSQL = cSQL & " and ADATE = " & lDatumJetzt
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.Edit
            Else
                rsrs.AddNew
            End If
            
            Select Case giZahlArt
                Case Is = 6     'Scheck
                    If Not IsNull(rsrs!TILGSCH) Then
                        rsrs!TILGSCH = rsrs!TILGSCH + dGPreis
                    Else
                        rsrs!TILGSCH = dGPreis
                    End If
                    
                Case Is = 8     'Bar
                    If Not IsNull(rsrs!TILGBAR) Then
                        rsrs!TILGBAR = rsrs!TILGBAR + dGPreis
                    Else
                        rsrs!TILGBAR = dGPreis
                    End If
                Case Is = 17    'Kreditkarte
                    If Not IsNull(rsrs!TILGKAR) Then
                        rsrs!TILGKAR = rsrs!TILGKAR + dGPreis
                        schreibeProtokollUNITXT CStr(dGPreis), "Kartenzahlung"
                    Else
                        rsrs!TILGKAR = dGPreis
                        schreibeProtokollUNITXT CStr(dGPreis), "Kartenzahlung"
                    End If
            End Select
            rsrs!KASNUM = gcKasNum
            rsrs!ADATE = lDatumJetzt
            rsrs.Update
            rsrs.Close: Set rsrs = Nothing
            
            cSQL = "Delete from KREDIT "
            cSQL = cSQL & "where KUNDNR = " & cKundnr & " "
            cSQL = cSQL & "and ADATE = " & Trim$(Str$(lDatum)) & " "
            cSQL = cSQL & "and ARTNR = " & cArtNr & " "
            cSQL = cSQL & "and MENGE = " & cAnzahl & " "
            schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
            
            'In die KreditProtokoll für die Zentrale
            
            lKJADate = Fix(Now)
            cKJAZeit = Format$(Now, "HH:MM:SS")
            
            insertKreditZA lKJADate, cKJAZeit, lDatum, CInt(gcBedienerNr), gcKreditKarte, CLng(cKundnr), CLng(cArtNr), CInt(cAnzahl), "0"
            
            'Kreditprotokoll ende
            
            MSFlexGrid2.Col = 0
            MSFlexGrid2.Text = "gelöscht"
            
            MSFlexGrid2.Col = 10
            MSFlexGrid2.Text = ""
            
            MSFlexGrid2.Col = 0
            
        End If
    Next lAktRecord

    giZahlArt = 0
    gcKreditKarte = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AusbuchenKreditVerkaufWKL24ohneBon"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub DruckeAlteKrediteWKL24()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim lRows       As Long
    Dim lrow        As Long
    Dim ctmp        As String
    
    loeschNEW "DRU_ALKR", gdBase
    CreateTable "DRU_ALKR", gdBase
    
    MSFlexGrid3.Redraw = False
    lRows = MSFlexGrid3.Rows
    
    cSQL = "Select * from DRU_ALKR where KUNDNR = -1"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    For lrow = 1 To lRows - 1
        MSFlexGrid3.Row = lrow
        
        rsrs.AddNew
        MSFlexGrid3.Col = 0
        ctmp = MSFlexGrid3.Text
        rsrs!REDATUM = Left(ctmp, 10)
        
        MSFlexGrid3.Col = 1
        ctmp = MSFlexGrid3.Text
        rsrs!RENR = ctmp

        MSFlexGrid3.Col = 2
        ctmp = MSFlexGrid3.Text
        rsrs!Kundnr = Val(ctmp)

        MSFlexGrid3.Col = 3
        ctmp = MSFlexGrid3.Text
        rsrs!name = Left(ctmp, 35)

        MSFlexGrid3.Col = 4
        ctmp = MSFlexGrid3.Text
        rsrs!strasse = Left(ctmp, 35)

        MSFlexGrid3.Col = 5
        ctmp = MSFlexGrid3.Text
        rsrs!Plz = Left(ctmp, 7)

        MSFlexGrid3.Col = 6
        ctmp = MSFlexGrid3.Text
        rsrs!Ort = Left(ctmp, 35)

        MSFlexGrid3.Col = 7
        ctmp = MSFlexGrid3.Text
        ctmp = fnMoveComma2Point(ctmp)
        rsrs!RESUMME = Val(ctmp)

        MSFlexGrid3.Col = 8
        ctmp = MSFlexGrid3.Text
        rsrs!Status = ctmp
    
        rsrs.Update
    
    Next lrow
    
    
    rsrs.Close: Set rsrs = Nothing
    MSFlexGrid3.Redraw = True
    reportbildschirm "WKL023a", "aWKL24"

    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeAlteKrediteWKL24"
        Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub DruckeOffeneKrediteWKL24()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lRows As Long
    Dim lrow As Long
    Dim lcol As Long
    Dim ctmp As String

    loeschNEW "DRU_OFKR", gdBase
    
    cSQL = "Create Table DRU_OFKR "
    cSQL = cSQL & "( KUNDNR Long"
    cSQL = cSQL & ", KUERZEL Text(5)"
    cSQL = cSQL & ", VORNAME Text(35)"
    cSQL = cSQL & ", NAME Text(35)"
    cSQL = cSQL & ", STRASSE Text(35)"
    cSQL = cSQL & ", PLZ Text(7)"
    cSQL = cSQL & ", ORT Text(35)"
    cSQL = cSQL & ", OFFEN Double"
    cSQL = cSQL & ", DATUM Text(10)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    lRows = MSFlexGrid1.Rows
    
    cSQL = "Select * from DRU_OFKR where KUNDNR = -1"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    For lrow = 1 To lRows - 1
        MSFlexGrid1.Row = lrow
        
        rsrs.AddNew
        MSFlexGrid1.Col = 0
        ctmp = MSFlexGrid1.Text
        rsrs!Kundnr = Val(ctmp)
        
        MSFlexGrid1.Col = 1
        ctmp = MSFlexGrid1.Text
        rsrs!Kuerzel = ctmp
        
        MSFlexGrid1.Col = 2
        ctmp = MSFlexGrid1.Text
        rsrs!vorname = ctmp
    
        MSFlexGrid1.Col = 3
        ctmp = MSFlexGrid1.Text
        rsrs!name = ctmp
    
        MSFlexGrid1.Col = 4
        ctmp = MSFlexGrid1.Text
        rsrs!strasse = ctmp
    
        MSFlexGrid1.Col = 5
        ctmp = MSFlexGrid1.Text
        rsrs!Plz = ctmp
    
        MSFlexGrid1.Col = 6
        ctmp = MSFlexGrid1.Text
        rsrs!Ort = ctmp
    
        MSFlexGrid1.Col = 7
        ctmp = MSFlexGrid1.Text
        ctmp = fnMoveComma2Point(ctmp)
        rsrs!offen = Val(ctmp)
        
        MSFlexGrid1.Col = 8
        ctmp = MSFlexGrid1.Text
        rsrs!Datum = ctmp
    
        rsrs.Update
    
    Next lrow
    rsrs.Close: Set rsrs = Nothing
    reportbildschirm "WKL023", "aWKL24a"

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "DruckeOffeneKrediteWKL24"
        Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Function fnPruefeStatusOffeneKrediteWKL24%()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim ctmp As String
    Dim bgefunden As Boolean
    
    fnPruefeStatusOffeneKrediteWKL24% = 0
    
    lAnzRecords = MSFlexGrid2.Rows - 1
    
    bgefunden = False
    MSFlexGrid2.Redraw = False
    If lAnzRecords > 0 Then
        
        For lAktRecord = 1 To lAnzRecords
            MSFlexGrid2.Row = lAktRecord
            MSFlexGrid2.Col = 0
            ctmp = MSFlexGrid2.Text
            ctmp = Trim$(ctmp)
            If ctmp = "ausbuchen" Then
                bgefunden = True
                Exit For
            End If
        Next lAktRecord
    End If
    MSFlexGrid2.Redraw = True
    
    If Not bgefunden Then
        MsgBox "Es wurden keine Daten für eine Rechnung gefunden!", vbInformation, "KEINE DATEN GEFUNDEN"
        fnPruefeStatusOffeneKrediteWKL24% = 1
    End If
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeStatusOffenekrediteWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub LeseKreditDetailsWKL24()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim dWert       As Double
    Dim lAktSatz    As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim iAdressID   As Integer
    Dim iWert       As Integer
    Dim aBreite(0 To 10) As Integer
    Dim iStufe      As Integer
    iStufe = 1
    
    bUnterschiedlicheRechnunganschriften = False
    
    
    With MSFlexGrid2
    
    .Cols = 11
    
    .Row = 0
    .Col = 0
    .Text = "Status"
    
    .Col = 1
    .Text = "Datum"
    
    .Col = 2
    .Text = "ArtNr"
    
    .Col = 3
    .Text = "Bezeich"
    
    .Col = 4
    .Text = "Anzahl"
    
    .Col = 5
    .Text = "Einzelpreis"
    
    .Col = 6
    .Text = "noch offen"
    
    .Col = 7
    .Text = "MWSt"
    
    .Col = 8
    .Text = "PreisKz"
    
    .Col = 9
    .Text = "Flag"
    
    .Col = 10
    .Text = "Reihenfolge"
    iStufe = 2
    For i = 0 To 10
        aBreite(i) = TextWidth(.TextMatrix(0, i))
    Next i
    
    iStufe = 3
    MSFlexGrid1.Row = Label0.Caption
    MSFlexGrid1.Col = 0
    Label2(0).Caption = MSFlexGrid1.Text
    iStufe = 4
    If Label2(0).Caption = "" Then
        Exit Sub
    End If
    iStufe = 5
    cSQL = "Select * from KUNDEN where KUNDNR = " & Label2(0).Caption
    Set rsrs = gdBase.OpenRecordset(cSQL)
    iStufe = 6
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        iStufe = 7
        Label2(6).Caption = IIf(IsNull(rsrs!vorname), "", rsrs!vorname)
        iStufe = 8
        Label2(1).Caption = IIf(IsNull(rsrs!name), "", rsrs!name)
        iStufe = 9
        cFirma = IIf(IsNull(rsrs!firma), "", rsrs!firma)
        iStufe = 10
        cAnrede = IIf(IsNull(rsrs!anrede), "", rsrs!anrede)
        iStufe = 11
        Label2(2).Caption = IIf(IsNull(rsrs!strasse), "", rsrs!strasse)
        iStufe = 12
        Label2(3).Caption = IIf(IsNull(rsrs!Plz), "", rsrs!Plz)
        iStufe = 13
        Label2(4).Caption = IIf(IsNull(rsrs!STADT), "", rsrs!STADT)
        iStufe = 14
        cTitel = IIf(IsNull(rsrs!titel), "", rsrs!titel)

    Else

    End If
    rsrs.Close: Set rsrs = Nothing
        
    MSFlexGrid1.Col = 7
    Label2(5).Caption = MSFlexGrid1.Text
    iStufe = 15
    cSQL = "Select * from KREDIT where KUNDNR = " & Label2(0).Caption & " order by ADATE, FLAG"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        .Rows = rsrs.RecordCount + 1
        iStufe = 16
        rsrs.MoveFirst
        lAktSatz = 0
        Do While Not rsrs.EOF
            lAktSatz = lAktSatz + 1
            iStufe = 17
            iWert = IIf(IsNull(rsrs!AdressID), 0, rsrs!AdressID)
            
            If lAktSatz > 1 Then
                iStufe = 18
                If iAdressID <> iWert Then
                    bUnterschiedlicheRechnunganschriften = True
                    iStufe = 19
                Else
                    iAdressID = iWert
                    bUnterschiedlicheRechnunganschriften = False
                    iStufe = 20
                End If
            Else
                iAdressID = iWert
                iStufe = 21
            End If
            
                iStufe = 22
                .Row = lAktSatz
                .Col = 0
                .Text = "offen"
                .Col = 1
                iStufe = 23
                .Text = IIf(IsNull(rsrs!ADATE), "", rsrs!ADATE)
                .Col = 2
                iStufe = 24
                .Text = IIf(IsNull(rsrs!artnr), "", rsrs!artnr)
                .Col = 3
                iStufe = 25
                .Text = IIf(IsNull(rsrs!BEZEICH), "", rsrs!BEZEICH)
                .Col = 4
                iStufe = 26
                .Text = IIf(IsNull(rsrs!MENGE), "", rsrs!MENGE)
                .Col = 5
                iStufe = 27
                'Achtung Änderung f. Schlanstein 050907
'                .Text = Format$(IIf(IsNull(rsrs!vkpr), "0", rsrs!vkpr), "###,##0.000")

                'zurückgeändert am 19.08.16
                .Text = Format$(IIf(IsNull(rsrs!vkpr), "0", rsrs!vkpr), "###,##0.00")
                .Col = 6
                iStufe = 28
                .Text = Format$(IIf(IsNull(rsrs!GVKPR), "0", rsrs!GVKPR), "###,##0.00")
                .Col = 7
                iStufe = 29
                .Text = IIf(IsNull(rsrs!MWST), "", rsrs!MWST)
                .Col = 8
                iStufe = 30
                .Text = Format$(IIf(IsNull(rsrs!PREISKZ), "0", rsrs!PREISKZ), "#0")
                .Col = 9
                iStufe = 31
                .Text = Format$(IIf(IsNull(rsrs!FLAG), "0", rsrs!FLAG), "####0")
            rsrs.MoveNext
        Loop
        .Col = 6
        dWert = 0
        iStufe = 32
        For i = 1 To rsrs.RecordCount
            .Row = i
            dWert = dWert + CDbl(.Text)
        Next
        Label2(5).Caption = CStr(dWert)
    End If
    rsrs.Close: Set rsrs = Nothing
    iStufe = 33
    
    TabellenbreiteanpassenNH MSFlexGrid2, 1.25 * gdTabfak
    
    iStufe = 34
        
    End With
    
    If bUnterschiedlicheRechnunganschriften = False Then
        iStufe = 35
        If iAdressID > 0 Then
            cSQL = "Select * from zadress where Adressid = " & iAdressID
            Set rsrs = gdBase.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                rsrs.MoveFirst
                iStufe = 36
                
                Text2(5).Text = IIf(IsNull(rsrs!firma), "", rsrs!firma)
                iStufe = 37
                Text2(6).Text = IIf(IsNull(rsrs!anrede), "", rsrs!anrede)
                iStufe = 38
                Text2(0).Text = IIf(IsNull(rsrs!name), "", rsrs!name)
                iStufe = 39
                Text2(1).Text = IIf(IsNull(rsrs!vorname), "", rsrs!vorname)
                iStufe = 40
                Text2(2).Text = IIf(IsNull(rsrs!strasse), "", rsrs!strasse)
                iStufe = 41
                Text2(3).Text = IIf(IsNull(rsrs!Plz), "", rsrs!Plz)
                iStufe = 42
                Text2(4).Text = IIf(IsNull(rsrs!Ort), "", rsrs!Ort)
                iStufe = 43
                Text2(7).Text = IIf(IsNull(rsrs!titel), "", rsrs!titel)
            End If
            rsrs.Close: Set rsrs = Nothing
        Else
        iStufe = 44
        Text2(0).Text = Label2(1).Caption
        Text2(1).Text = Label2(6).Caption
        Text2(2).Text = Label2(2).Caption
        Text2(3).Text = Label2(3).Caption
        Text2(4).Text = Label2(4).Caption
        Text2(5).Text = cFirma
        Text2(6).Text = cAnrede
        Text2(7).Text = cTitel
        
        iStufe = 45
        End If
        
    Else
        iStufe = 46
        lblUeber(4).Caption = "Achtung unterschiedliche Rechnungsanschriften!!!"
    End If
    
    Label5(0).Caption = ""
    Label5(1).Caption = ""
    Label5(2).Caption = ""
    
    Frame2.Visible = True
    Frame1.Visible = False
    
    Command2(0).Enabled = True
    Check3.Visible = True
    
    iStufe = 47
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LesekreditDetailsWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten. " & iStufe & " " & Label0.Caption
    
    Fehlermeldung1
    
End Sub
Private Sub LeseOffeneKrediteWKL24(iTage As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lAktSatz As Long
    Dim cKdnr As String
    Dim dOffen As Double
    Dim ctmp As String
    Dim lDatum As Long
    Dim cDatum As String
    Dim i As Integer
    Dim j As Integer
    Dim aBreite(0 To 8) As Integer
    
    picprogress.Visible = True
    txtStatus.Text = "13"
    
    loeschNEW "KRA", gdBase
    txtStatus.Text = "15"
    CreateTable "KRA", gdBase
    
    txtStatus.Text = "26"
    
    MSFlexGrid1.Clear
    
    With MSFlexGrid1
    
        .Redraw = False
        
        .Row = 0
        .Col = 0
        .Text = "KundNr"
        .Col = 1
        .Text = "Kürzel"
        .Col = 2
        .Text = "Vorname"
        .Col = 3
        .Text = "Name"
        .Col = 4
        .Text = "Straße"
        .Col = 5
        .Text = "PLZ"
        .Col = 6
        .Text = "Ort"
        .Col = 7
        .Text = "offen"
        .Col = 8
        .Text = "Datum"
        
        For i = 0 To 8
            aBreite(i) = TextWidth(.TextMatrix(0, i))
        Next i
        txtStatus.Text = "34"
        loeschNEW "KR1A", gdBase
        txtStatus.Text = "45"
        
        cSQL = "Select KUNDNR, SUM(GVKPR) as OFFEN, ADATE into Kr1a from KREDIT "
        If iTage <> 0 Then
            lblUeber(1).Caption = "Kunden mit offenen Krediten der letzten 61 Tage"
            cSQL = cSQL & " where adate > Datevalue(now) - " & iTage
        Else
            lblUeber(1).Caption = "Kunden mit offenen Krediten alle"
        End If
        cSQL = cSQL & " group by KUNDNR, ADATE order by KUNDNR"
        gdBase.Execute cSQL, dbFailOnError
        
        txtStatus.Text = "58"
        
        cSQL = "Insert into KRA select * from KR1A"
        gdBase.Execute cSQL, dbFailOnError
        
        txtStatus.Text = "63"
        
        cSQL = "Update KRA inner join Kunden on KRA.KUNDNR = Kunden.KUNDNR "
        cSQL = cSQL & " Set kra.name = kunden.name "
        cSQL = cSQL & " , kra.Vorname = kunden.vorname "
        cSQL = cSQL & " , kra.PLZ = kunden.PLZ "
        cSQL = cSQL & " , kra.STADT = kunden.STADT "
        cSQL = cSQL & " , kra.STRASSE = kunden.STRASSE "
        cSQL = cSQL & " , kra.KUERZEL = kunden.KUERZEL "
        gdBase.Execute cSQL, dbFailOnError
        
        
        txtStatus.Text = "81"
        
        cSQL = "Select * from  Kra"
        FnOpenrecordset rsrs, cSQL, 1, gdBase
        
        txtStatus.Text = "99"
        
        If Not rsrs.EOF Then
            rsrs.MoveLast
            .Rows = rsrs.RecordCount + 1
            rsrs.MoveFirst
            lAktSatz = 0
            Do While Not rsrs.EOF
                lAktSatz = lAktSatz + 1
                txtStatus.Text = (lAktSatz * 100) / .Rows
                
                .Row = lAktSatz
                .Col = 0
                
                If Not IsNull(rsrs!Kundnr) Then
                    cKdnr = rsrs!Kundnr
                Else
                    cKdnr = "0"
                End If
                .Text = cKdnr
                
                .Col = 1
                If Not IsNull(rsrs!Kuerzel) Then
                    .Text = rsrs!Kuerzel
                Else
                    .Text = ""
                End If
                
                .Col = 2
                If Not IsNull(rsrs!vorname) Then
                    .Text = rsrs!vorname
                Else
                    .Text = ""
                End If
                    
                .Col = 3
                If Not IsNull(rsrs!name) Then
                    .Text = rsrs!name
                Else
                    .Text = ""
                End If
                .Col = 4
                
                If Not IsNull(rsrs!strasse) Then
                    .Text = rsrs!strasse
                Else
                    .Text = ""
                End If
                   
                .Col = 5
                If Not IsNull(rsrs!Plz) Then
                    .Text = rsrs!Plz
                Else
                    .Text = ""
                End If
                .Col = 6
                If Not IsNull(rsrs!STADT) Then
                    .Text = rsrs!STADT
                Else
                    .Text = ""
                End If
                
                .Col = 7
                If Not IsNull(rsrs!offen) Then
                    .Text = Format$(rsrs!offen, "######0.00")
                Else
                    .Text = "0"
                End If
                
                .Col = 8
                If Not IsNull(rsrs!ADATE) Then
                    .Text = Format$(rsrs!ADATE, "DD.MM.YYYY")
                Else
                    .Text = ""
                End If
                
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
    
        
        TabellenbreiteanpassenNH MSFlexGrid1, 1.25 * gdTabfak
        .Redraw = True
    End With
    
    If MSFlexGrid1.Rows > 1 Then
'        MSFlexGrid1.SetFocus
        MSFlexGrid1.Row = 1
        MSFlexGrid1.Col = 2
    End If
    
    picprogress.Visible = False
    txtStatus.Text = "0"
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseOffenekrediteWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MoveDaten2RechnungWKL24(lAktRecord As Long, cReNr As String, dSumme As Double, lFlag As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim dWert As Double
    Dim lDatum As Long
    Dim cDatum As String
    Dim cArtNr As String
    Dim cBezeich As String
    Dim cAnzahl As String
    Dim cEPreis As String
    Dim cGPreis As String
    Dim cMwst As String
    Dim cKdnr As String
    Dim cPreisKz As String
    
    Dim cInto As String
    Dim cSQL As String
    
    cKdnr = Label2(0).Caption
    
    MSFlexGrid2.Row = lAktRecord
        
    MSFlexGrid2.Col = 1
    ctmp = MSFlexGrid2.Text
    lDatum = DateValue(ctmp)
    cDatum = Trim$(Str$(lDatum))
    
    MSFlexGrid2.Col = 2
    cArtNr = MSFlexGrid2.Text
        
    MSFlexGrid2.Col = 3
    cBezeich = MSFlexGrid2.Text
        
    MSFlexGrid2.Col = 4
    ctmp = MSFlexGrid2.Text
    ctmp = fnMoveComma2Point$(ctmp)
    cAnzahl = ctmp
    
    MSFlexGrid2.Col = 5
    ctmp = MSFlexGrid2.Text
    If ctmp = "" Then
        ctmp = "0"
    End If
    ctmp = fnMoveComma2Point$(ctmp)
    cEPreis = ctmp
        
    MSFlexGrid2.Col = 6
    ctmp = MSFlexGrid2.Text
    ctmp = fnMoveComma2Point$(ctmp)
    cGPreis = ctmp
    dSumme = Val(ctmp)
    
    MSFlexGrid2.Col = 7
    ctmp = MSFlexGrid2.Text
    ctmp = fnMoveComma2Point$(ctmp)
    cMwst = ctmp
    
    MSFlexGrid2.Col = 8
    ctmp = MSFlexGrid2.Text
    ctmp = fnMoveComma2Point$(ctmp)
    cPreisKz = ctmp
    
    Dim lreihenfolge As Long
    
    MSFlexGrid2.Col = 10
    ctmp = MSFlexGrid2.Text
    ctmp = fnMoveComma2Point$(ctmp)
    lreihenfolge = CLng(ctmp)
    
    
    cInto = "Insert Into REPOS "
    cSQL = "("
    cSQL = cSQL & " SCHLUESSEL"
    cSQL = cSQL & ", KAUFDATUM"
    cSQL = cSQL & ", ARTNR"
    cSQL = cSQL & ", BEZEICH"
    cSQL = cSQL & ", ANZAHL"
    cSQL = cSQL & ", EPREIS"
    cSQL = cSQL & ", GPREIS"
    cSQL = cSQL & ", MWST"
    cSQL = cSQL & ", PREISKZ"
    cSQL = cSQL & ", Reihenf "
    cSQL = cSQL & ") values ("
    cSQL = cSQL & "'" & cReNr & "'"
    cSQL = cSQL & ", " & cDatum & ""
    cSQL = cSQL & ", " & cArtNr & ""
    cSQL = cSQL & ", '" & cBezeich & "' "
    cSQL = cSQL & ", " & cAnzahl & ""
    cSQL = cSQL & ", " & cEPreis & ""
    cSQL = cSQL & ", " & cGPreis & ""
    cSQL = cSQL & ", '" & cMwst & "' "
    cSQL = cSQL & ", " & cPreisKz & ""
    cSQL = cSQL & ", " & lreihenfolge & ""
    cSQL = cSQL & ") "
    gdBase.Execute cInto & cSQL, dbFailOnError

    cSQL = "Delete from KREDIT where "
    cSQL = cSQL & "ARTNR = " & cArtNr & " and "
    cSQL = cSQL & "MENGE = " & cAnzahl & " and "
    cSQL = cSQL & "KUNDNR = " & cKdnr & " and "
    cSQL = cSQL & "ADATE = " & cDatum & " and "
    cSQL = cSQL & "FLAG = " & Trim$(Str$(lFlag)) & " "
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveDaten2RechnungWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MoveRePos2DruRePosWKL24(cReNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cNettoVoll  As String
    Dim cNettoErm   As String
    
    Screen.MousePointer = 11
    
    loeschNEW "DRU_REPO", gdBase
    
    
    
    cSQL = "Create Table DRU_REPO "
    cSQL = cSQL & "(SCHLUESSEL Text(15)"
    cSQL = cSQL & ", KAUFDATUM Datetime"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH Text(100)"
    cSQL = cSQL & ", ANZAHL long"
    cSQL = cSQL & ", EPREIS double"
    cSQL = cSQL & ", GPREIS double"
    cSQL = cSQL & ", MWST Text(1)"
    cSQL = cSQL & ", PREISKZ Integer"
    cSQL = cSQL & ", STEUERNR Text(35)"
    cSQL = cSQL & ", Reihenf long"
    cSQL = cSQL & ", FILIALE Integer"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert Into DRU_REPO "
    cSQL = cSQL & "Select * from REPOS "
    cSQL = cSQL & " where SCHLUESSEL = '" & cReNr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    cNettoErm = Format(ermNettoERM, "#####0.00")
    cNettoErm = SwapStr(cNettoErm, ",", ".")
    
'    cNettoVoll = ermNettoVoll
    cNettoVoll = Format(ermNettoVoll, "#####0.00")
    cNettoVoll = SwapStr(cNettoVoll, ",", ".")
    
    loeschNEW "NETTOS", gdBase
    
    cSQL = "Create Table NETTOS "
    cSQL = cSQL & "("
    cSQL = cSQL & " NettERM double"
    cSQL = cSQL & ", NettVol double"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert Into NETTOS (NETTERM , NETTVOL) values ( " & cNettoErm & "," & cNettoVoll & " ) "
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MoveRePos2DruRePosWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Function ermNettoERM() As String
    On Error GoTo LOKAL_ERROR
    
    ermNettoERM = "0"
    
    Dim cSQL As String
    Dim rs As Recordset
    
    cSQL = "Select Sum(GPreis)as maxi from DRU_REPO where MWST = 'E' "
    Set rs = gdBase.OpenRecordset(cSQL)
    
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            ermNettoERM = rs!maxi
            ermNettoERM = (ermNettoERM * 100) / (100 + gdMWStE)
        End If
    End If
    rs.Close: Set rs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermNettoERM"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermNettoVoll() As String
    On Error GoTo LOKAL_ERROR
    
    ermNettoVoll = "0"
    
    Dim cSQL As String
    Dim rs As Recordset
    
    cSQL = "Select Sum(GPreis)as maxi from DRU_REPO where MWST = 'V' "
    Set rs = gdBase.OpenRecordset(cSQL)
    
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            ermNettoVoll = rs!maxi
            ermNettoVoll = (ermNettoVoll * 100) / (100 + gdMWStV)
        End If
    End If
   
    rs.Close: Set rs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermNettoVoll"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Private Sub RechnungRueckgaengigWKL24()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim cSchluessel As String
    Dim cKundnr As String
    Dim cSQL As String
    
    lrow = Val(Label4.Caption)
    
    MSFlexGrid3.Row = lrow
    MSFlexGrid3.Col = 1
    cSchluessel = MSFlexGrid3.Text
    
    MSFlexGrid3.Row = lrow
    MSFlexGrid3.Col = 2
    cKundnr = MSFlexGrid3.Text
    
    cSQL = "Insert into KREDIT "
    cSQL = cSQL & "Select -99 as FLAG, ARTNR, ANZAHL as MENGE, EPREIS as VKPR"
    cSQL = cSQL & ", PREISKZ, MWST, BEZEICH, KAUFDATUM as ADATE "
    cSQL = cSQL & ", " & cKundnr & " as KUNDNR "
    cSQL = cSQL & "from REPOS where SCHLUESSEL = '" & cSchluessel & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KREDIT inner join ARTIKEL on KREDIT.ARTNR = ARTIKEL.ARTNR "
    cSQL = cSQL & "Set KREDIT.EKPR = ARTIKEL.EKPR, KREDIT.AVKPR = ARTIKEL.KVKPR1 "
    cSQL = cSQL & "where KREDIT.FLAG = -99 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update KREDIT Set FLAG = 0, GVKPR = MENGE * VKPR where FLAG = -99"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "DELETE from REPOS where SCHLUESSEL = '" & cSchluessel & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "DELETE from REKOPF where SCHLUESSEL = '" & cSchluessel & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    'auch OFPO
    cSQL = "DELETE from OFPO where SCHLUESSEL = '" & cSchluessel & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    If Check2.value = vbChecked Then
        LeseOffeneKrediteWKL24 0
    Else
        LeseOffeneKrediteWKL24 61
    End If
'    LeseOffeneKrediteWKL24 0
    
    ZeigeVorhandeneRechnungenWKL24
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "RechnungRueckgaengigWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub RechnungLoeschenWKL24()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim cSchluessel As String
    Dim cKundnr As String
    Dim cSQL As String
    Dim cDatum As String
    Dim lDatum As Long
    Dim iRet As Integer
    Dim ctmp As String
    Dim bTrans As Boolean
    Dim rsKopf As Recordset
    Dim rsPos As Recordset
    
    bTrans = False
    
    lrow = Val(Label4.Caption)
    
    
    cSchluessel = MSFlexGrid3.TextMatrix(lrow, 1)
    
    
    If cSchluessel = "" Then
        Exit Sub
    End If
    
    cDatum = MSFlexGrid3.TextMatrix(lrow, 0)
    lDatum = DateValue(cDatum)
    
    
    
    cKundnr = MSFlexGrid3.TextMatrix(lrow, 2)
    
    ctmp = "Wollen Sie die Rechnung" & vbCrLf & vbCrLf
    ctmp = ctmp & "vom     " & cDatum & vbCrLf
    ctmp = ctmp & "RechNr. " & cSchluessel & vbCrLf
    ctmp = ctmp & "KundNr. " & cKundnr & vbCrLf & vbCrLf
    ctmp = ctmp & "wirklich löschen?"

    iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "RECHNUNG LÖSCHEN")
    If iRet <> vbYes Then
        Exit Sub
    End If
    
    cSQL = "Select * from REKOPF "
    cSQL = cSQL & "where SCHLUESSEL = '" & cSchluessel & "' "
    cSQL = cSQL & "and KUNDNR = " & cKundnr & " "
    cSQL = cSQL & "and REDATUM = " & Trim$(Str$(lDatum)) & " "
    Set rsKopf = gdBase.OpenRecordset(cSQL)
    
    cSQL = "Select * from REPOS where SCHLUESSEL = '" & cSchluessel & "' "
    Set rsPos = gdBase.OpenRecordset(cSQL)
    
    BeginTrans
    bTrans = True
    
    If Not rsPos.EOF Then
        rsPos.MoveFirst
        Do While Not rsPos.EOF
            rsPos.delete
            rsPos.MoveNext
        Loop
    End If
    
    If Not rsKopf.EOF Then
        rsKopf.MoveFirst
        Do While Not rsKopf.EOF
            rsKopf.delete
            rsKopf.MoveNext
        Loop
    End If
    CommitTrans
    bTrans = False
    
    rsPos.Close
    rsKopf.Close
    
    If Check2.value = vbChecked Then
        LeseOffeneKrediteWKL24 0
    Else
        LeseOffeneKrediteWKL24 61
    End If
'    LeseOffeneKrediteWKL24 0
    
    ZeigeVorhandeneRechnungenWKL24
    
Exit Sub
LOKAL_ERROR:
    If bTrans Then
        Rollback
    End If
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "RechnungLoeschenWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
'    Resume Next
End Sub
Private Sub SchreibeAltRechnungWKL24(bReFuss As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim cReNr As String
    Dim cDatum As String
    Dim cSQL            As String
    Dim lDatum          As Long
    Dim rsrs            As Recordset
    Dim rsF             As Recordset
    Dim cRePreisKz      As String
    
    Dim cFirmName       As String
    Dim cFirmAdress     As String
    Dim cFirmBank       As String
    Dim cFirmKomm       As String
    Dim cSteuernr       As String
    
    Dim cMWStKz As String
    Dim iMixMWSt As Integer
    
    MSFlexGrid3.Row = Val(Label4.Caption)
    
    MSFlexGrid3.Col = 1
    cReNr = MSFlexGrid3.Text
    
    If cReNr = "" Then
        Exit Sub
    End If
    
    If bReFuss Then
        cSQL = "Select * from FIRMA"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Steuernr) Then
                cSteuernr = rsrs!Steuernr
            Else
                cSteuernr = ""
            End If
            If Not IsNull(rsrs!name) Then
                cFirmName = rsrs!name
            Else
                cFirmName = ""
            End If
            If Not IsNull(rsrs!strasse) Then
                cFirmAdress = rsrs!strasse
            Else
                cFirmAdress = ""
            End If
            If Not IsNull(rsrs!Plz) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & "   " & rsrs!Plz
                Else
                    cFirmAdress = rsrs!Plz
                End If
            End If
            If Not IsNull(rsrs!Ort) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & " " & rsrs!Ort
                Else
                    cFirmAdress = rsrs!Ort
                End If
            End If
            If Not IsNull(rsrs!BankName) Then
                cFirmBank = rsrs!BankName
            Else
                cFirmBank = ""
            End If
            If Not IsNull(rsrs!BLZ) Then
                If rsrs!BLZ <> "" Then
                    cFirmBank = cFirmBank & "  BLZ " & rsrs!BLZ
                End If
            End If
            
            If Not IsNull(rsrs!Konto) Then
                If rsrs!Konto <> "" Then
                    cFirmBank = cFirmBank & "  Konto: " & rsrs!Konto
                End If
            End If
            
            If Not IsNull(rsrs!BIC) Then
                If rsrs!BIC <> "" Then
                    cFirmBank = cFirmBank & "  BIC " & rsrs!BIC
                End If
            End If
            
            If Not IsNull(rsrs!IBAN) Then
                If rsrs!IBAN <> "" Then
                    cFirmBank = cFirmBank & "  IBAN: " & rsrs!IBAN
                End If
            End If
            If Not IsNull(rsrs!Tel) Then
                cFirmKomm = "Tel.: " & rsrs!Tel
            Else
                cFirmKomm = ""
            End If
            If Not IsNull(rsrs!Fax) Then
                If cFirmKomm <> "" Then
                    cFirmKomm = cFirmKomm & "  Fax: " & rsrs!Fax
                Else
                    cFirmKomm = "Fax: " & rsrs!Fax
                End If
            End If
        Else
            cSteuernr = ""
            cFirmName = ""
            cFirmAdress = ""
            cFirmBank = ""
            cFirmKomm = ""
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
    MSFlexGrid3.Row = Val(Label4.Caption)
    
    MSFlexGrid3.Col = 1
    cReNr = MSFlexGrid3.Text
    
    
    
    MSFlexGrid3.Col = 0
    cDatum = MSFlexGrid3.Text
        
    lDatum = DateValue(cDatum)
    cDatum = Trim$(Str$(lDatum))
    
    
    
    
    'Druck-Datei löschen und neu erzeugen

    loeschNEW "DRU_REKO", gdBase
    
    cSQL = "Create Table DRU_REKO "
    cSQL = cSQL & "(SCHLUESSEL text(15)"
    cSQL = cSQL & ", KUNDNR LONG"
    cSQL = cSQL & ", ANREDE text(35)"
    cSQL = cSQL & ", KDNAME1 text(71)"
    cSQL = cSQL & ", KDNAME2 text(71)"
    cSQL = cSQL & ", STRASSE text(35)"
    cSQL = cSQL & ", PLZ text(7)"
    cSQL = cSQL & ", ORT text(35)"
    cSQL = cSQL & ", RENR Text(15)"
    cSQL = cSQL & ", REDATUM Datetime"
    cSQL = cSQL & ", RETEXT text(100)"
    cSQL = cSQL & ", RESUMME double"
    cSQL = cSQL & ", STATUS text(1)"
    cSQL = cSQL & ", PORTOVERP double"
    cSQL = cSQL & ", KOMMENTAR memo"
    cSQL = cSQL & ", PREISKZ Text(1)"
    cSQL = cSQL & ", STEUERNR Text(35)"
    cSQL = cSQL & ", FIRMNAME Text(50)"
    cSQL = cSQL & ", FIRMADRESS Text(110)"
    cSQL = cSQL & ", FIRMBANK Text(200)"
    cSQL = cSQL & ", FIRMKOMM Text(100)"
    cSQL = cSQL & ", IHR_ZEICHEN Text(30)"
    cSQL = cSQL & ", ERSTELLT_VON Text(30)"
    cSQL = cSQL & ", ZAHLUNGS_ZIEL Text(250)"
    cSQL = cSQL & ", SPEZIALSATZ Text(250)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "DRU_REPO", gdBase
    
    cSQL = "Create Table DRU_REPO "
    cSQL = cSQL & "(SCHLUESSEL text(15)"
    cSQL = cSQL & ", KAUFDATUM Datetime"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH text(100)"
    cSQL = cSQL & ", ANZAHL long"
    cSQL = cSQL & ", EPREIS double"
    cSQL = cSQL & ", GPREIS double"
    cSQL = cSQL & ", MWST text(1)"
    cSQL = cSQL & ", PREISKZ Integer"
    cSQL = cSQL & ", STEUERNR Text(35)"
    cSQL = cSQL & ", Reihenf long"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    Dim cNettoVoll  As String
    Dim cNettoErm   As String
    
    cSQL = "Insert into DRU_REKO "
    cSQL = cSQL & "Select * "
    cSQL = cSQL & ", '" & cSteuernr & "' as STEUERNR"
    cSQL = cSQL & ", '" & cFirmName & "' as FIRMNAME"
    cSQL = cSQL & ", '" & cFirmAdress & "' as FIRMADRESS"
    cSQL = cSQL & ", '" & cFirmBank & "' as FIRMBANK"
    cSQL = cSQL & ", '" & cFirmKomm & "' as FIRMKOMM"
    cSQL = cSQL & " from REKOPF where SCHLUESSEL = '" & cReNr & "' and REDATUM = " & cDatum & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into DRU_REPO Select * "
    cSQL = cSQL & "  from REPOS where SCHLUESSEL = '" & cReNr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cNettoErm = Format(ermNettoERM, "#####0.00")
    cNettoErm = SwapStr(cNettoErm, ",", ".")
    
    cNettoVoll = Format(ermNettoVoll, "#####0.00")
    cNettoVoll = SwapStr(cNettoVoll, ",", ".")
    
    loeschNEW "NETTOS", gdBase
    
    cSQL = "Create Table NETTOS "
    cSQL = cSQL & "("
    cSQL = cSQL & " NettERM double"
    cSQL = cSQL & ", NettVol double"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert Into NETTOS (NETTERM , NETTVOL) values ( " & cNettoErm & "," & cNettoVoll & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index SCHLUESSEL from DRU_REKO"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index SCHLUESSEL from DRU_REPO"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index SCHLUESSEL on DRU_REKO (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index SCHLUESSEL on DRU_REPO (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from DRU_REKO"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!PREISKZ) Then
            cRePreisKz = rsrs!PREISKZ
        Else
            cRePreisKz = "N"
        End If
    Else
        cRePreisKz = "N"
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from FIRMA"
    Set rsF = gdBase.OpenRecordset(cSQL)
    If Not rsF.EOF Then
        If Not IsNull(rsF!Steuernr) Then
            cSteuernr = rsF!Steuernr
        Else
            cSteuernr = "keine Angabe"
        End If
    End If
    rsF.Close
    
    cSQL = "Update DRU_REPO SET STEUERNR = '" & cSteuernr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from DRU_REPO"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    iMixMWSt = 0
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!MWST) Then
                cMWStKz = rsrs!MWST
            Else
                cMWStKz = ""
            End If
            Select Case cMWStKz
                Case Is = "V"
                    If Not iMixMWSt And 1 Then
                        iMixMWSt = iMixMWSt + 1
                    End If
                Case Is = "E"
                    If Not iMixMWSt And 2 Then
                        iMixMWSt = iMixMWSt + 2
                    End If
                Case Is = "O"
                    If Not iMixMWSt And 4 Then
                        iMixMWSt = iMixMWSt + 4
                    End If
                Case Else
                    'ignorieren
            End Select
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Select Case iMixMWSt
        Case 0, 1, 2, 4     'MWST-Kz nur V oder nur E oder nur O oder leer
            Select Case cRePreisKz
                Case Is = "N"
                    If Modul6.FindFile(gcDBPfad, "aWKL24l.rpt") Then
                        reportbildschirm "spez5", "aWKL24l"
                    Else
                        reportbildschirm "WKL005", "aWKL24b"
                    End If
                Case Is = "B"
                    If Modul6.FindFile(gcDBPfad, "aWKL24m.rpt") Then
                    
                    
                        'bau mal eine Tabelle für die regulären KVK und ArtRab in Proz
    
                        If Modul6.FindFile(gcDBPfad, "aWKL24m.rpt") And UCase(gsStammFTPUSER) = "HICKMANN" Then
                             
                            cSQL = "Alter Table DRU_REPO add ARTRAB Double "
                            gdBase.Execute cSQL, dbFailOnError
                            
                            cSQL = "Alter Table DRU_REPO add KVKPR1 Double "
                            gdBase.Execute cSQL, dbFailOnError
                            
                            cSQL = "Update DRU_REPO inner join Artikel on DRU_REPO.Artnr = Artikel.Artnr "
                            cSQL = cSQL & " Set DRU_REPO.KVKPR1 = Artikel.KVKPR1 "
                            gdBase.Execute cSQL, dbFailOnError
                            
                            cSQL = "Update DRU_REPO SET Artrab = 100 - (EPREIS*100/kvkpr1) where kvkpr1 <> 0  "
                            gdBase.Execute cSQL, dbFailOnError
                            
                            cSQL = "Update DRU_REPO SET Artrab = Artrab * (-1)  "
                            gdBase.Execute cSQL, dbFailOnError
                            
                        End If
                        
                        'Ende, bau mal eine Tabelle
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                        reportbildschirm "spez5a", "aWKL24m"
                    Else
                        reportbildschirm "WKL005a", "aWKL24c"
                    End If
                Case Is = "O"
                    If Modul6.FindFile(gcDBPfad, "aWKL24n.rpt") Then
                        reportbildschirm "spez5b", "aWKL24n"
                    Else
                        reportbildschirm "WKL005b", "aWKL24d"
                    End If
                Case Is = "Z"
                    If Modul6.FindFile(gcDBPfad, "aWKL24o.rpt") Then
                        reportbildschirm "spez5c", "aWKL24o"
                    Else
                        reportbildschirm "WKL005c", "aWKL24e"
                    End If
            End Select
        
        Case 3, 5, 6, 7         'MWSt-Kz ist gemischt
            Select Case cRePreisKz
                Case Is = "N"
                    If Modul6.FindFile(gcDBPfad, "aWKL24l.rpt") Then
                        reportbildschirm "spez5d", "aWKL24l"
                    Else
                        reportbildschirm "WKL005d", "awkl24b"
                    End If
                Case Is = "B"
                    If Modul6.FindFile(gcDBPfad, "aWKL24m.rpt") Then
                        reportbildschirm "spez5e", "aWKL24m"
                    Else
                        reportbildschirm "WKL005e", "awkl24c"
                    End If
                Case Is = "O"
                    If Modul6.FindFile(gcDBPfad, "aWKL24n.rpt") Then
                        reportbildschirm "spez5f", "aWKL24n"
                    Else
                        reportbildschirm "WKL005f", "awkl24d"
                    End If
                Case Is = "Z"
                    If Modul6.FindFile(gcDBPfad, "aWKL24o.rpt") Then
                        reportbildschirm "spez5g", "aWKL24o"
                    Else
                        reportbildschirm "WKL005g", "awkl24e"
                    End If
                    
            End Select
    End Select
    

Exit Sub
LOKAL_ERROR:
    If err.Number = 3295 Or err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeAltRechnungWKL24"
        Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Public Sub SchreibeMahnung(iReFuss As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cReNr As String
    Dim cDatum As String
    Dim cSQL            As String
    Dim lDatum          As Long
    Dim rsrs            As Recordset
    Dim rsF             As Recordset
    Dim cRePreisKz      As String
    
    Dim cFirmName       As String
    Dim cFirmAdress     As String
    Dim cFirmBank       As String
    Dim cFirmKomm       As String
    Dim cSteuernr       As String
    Dim cKommentar      As String
    
    Dim cMWStKz As String
    Dim iMixMWSt As Integer
    
    If iReFuss = vbYes Then
        cSQL = "Select * from FIRMA"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Steuernr) Then
                cSteuernr = rsrs!Steuernr
            Else
                cSteuernr = ""
            End If
            If Not IsNull(rsrs!name) Then
                cFirmName = rsrs!name
            Else
                cFirmName = ""
            End If
            If Not IsNull(rsrs!strasse) Then
                cFirmAdress = rsrs!strasse
            Else
                cFirmAdress = ""
            End If
            If Not IsNull(rsrs!Plz) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & "   " & rsrs!Plz
                Else
                    cFirmAdress = rsrs!Plz
                End If
            End If
            If Not IsNull(rsrs!Ort) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & " " & rsrs!Ort
                Else
                    cFirmAdress = rsrs!Ort
                End If
            End If
            If Not IsNull(rsrs!BankName) Then
                cFirmBank = rsrs!BankName
            Else
                cFirmBank = ""
            End If
            If Not IsNull(rsrs!BLZ) Then
                If rsrs!BLZ <> "" Then
                    cFirmBank = cFirmBank & "  BLZ " & rsrs!BLZ
                End If
            End If
            
            If Not IsNull(rsrs!Konto) Then
                If rsrs!Konto <> "" Then
                    cFirmBank = cFirmBank & "  Konto: " & rsrs!Konto
                End If
            End If
            
            If Not IsNull(rsrs!BIC) Then
                If rsrs!BIC <> "" Then
                    cFirmBank = cFirmBank & "  BIC " & rsrs!BIC
                End If
            End If
            
            If Not IsNull(rsrs!IBAN) Then
                If rsrs!IBAN <> "" Then
                    cFirmBank = cFirmBank & "  IBAN: " & rsrs!IBAN
                End If
            End If
            If Not IsNull(rsrs!Tel) Then
                cFirmKomm = "Tel.: " & rsrs!Tel
            Else
                cFirmKomm = ""
            End If
            If Not IsNull(rsrs!Fax) Then
                If cFirmKomm <> "" Then
                    cFirmKomm = cFirmKomm & "  Fax: " & rsrs!Fax
                Else
                    cFirmKomm = "Fax: " & rsrs!Fax
                End If
            End If
        Else
            cSteuernr = ""
            cFirmName = ""
            cFirmAdress = ""
            cFirmBank = ""
            cFirmKomm = ""
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
    MSFlexGrid3.Row = Val(Label4.Caption)
    MSFlexGrid3.Col = 0
    cDatum = MSFlexGrid3.Text
        
    lDatum = DateValue(cDatum)
    cDatum = Trim$(Str$(lDatum))
    
    MSFlexGrid3.Col = 1
    cReNr = MSFlexGrid3.Text
    
    cKommentar = Label5(1).Caption
    
    
    'Druck-Datei löschen und neu erzeugen

    loeschNEW "DRU_REKO", gdBase
    
    cSQL = "Create Table DRU_REKO "
    cSQL = cSQL & "(SCHLUESSEL text(15)"
    cSQL = cSQL & ", KUNDNR LONG"
    cSQL = cSQL & ", ANREDE text(35)"
    cSQL = cSQL & ", KDNAME1 text(71)"
    cSQL = cSQL & ", KDNAME2 text(71)"
    cSQL = cSQL & ", STRASSE text(35)"
    cSQL = cSQL & ", PLZ text(7)"
    cSQL = cSQL & ", ORT text(35)"
    cSQL = cSQL & ", RENR Text(15)"
    cSQL = cSQL & ", REDATUM Datetime"
    cSQL = cSQL & ", RETEXT text(100)"
    cSQL = cSQL & ", RESUMME double"
    cSQL = cSQL & ", STATUS text(1)"
    cSQL = cSQL & ", PORTOVERP double"
    cSQL = cSQL & ", KOMMENTAR memo " ' text(250)"
    cSQL = cSQL & ", PREISKZ Text(1)"
    cSQL = cSQL & ", STEUERNR Text(35)"
    cSQL = cSQL & ", FIRMNAME Text(50)"
    cSQL = cSQL & ", FIRMADRESS Text(110)"
    cSQL = cSQL & ", FIRMBANK Text(100)"
    cSQL = cSQL & ", FIRMKOMM Text(100)"
    
    cSQL = cSQL & ", IHR_ZEICHEN Text(30)"
    cSQL = cSQL & ", ERSTELLT_VON Text(30)"
    cSQL = cSQL & ", ZAHLUNGS_ZIEL Text(250)"
    cSQL = cSQL & ", SPEZIALSATZ Text(250)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError

    loeschNEW "DRU_REPO", gdBase
    cSQL = "Create Table DRU_REPO "
    cSQL = cSQL & "(SCHLUESSEL text(15)"
    cSQL = cSQL & ", KAUFDATUM Datetime"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", BEZEICH text(35)"
    cSQL = cSQL & ", ANZAHL long"
    cSQL = cSQL & ", EPREIS double"
    cSQL = cSQL & ", GPREIS double"
    cSQL = cSQL & ", MWST text(1)"
    cSQL = cSQL & ", PREISKZ Integer"
    cSQL = cSQL & ", STEUERNR Text(35)"
    cSQL = cSQL & ", Reihenf long"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    Dim cNettoVoll  As String
    Dim cNettoErm   As String
    
    cSQL = "Insert into DRU_REKO "
    cSQL = cSQL & "Select * "
    cSQL = cSQL & ", '" & cSteuernr & "' as STEUERNR"
    cSQL = cSQL & ", '" & cFirmName & "' as FIRMNAME"
    cSQL = cSQL & ", '" & cFirmAdress & "' as FIRMADRESS"
    cSQL = cSQL & ", '" & cFirmBank & "' as FIRMBANK"
    cSQL = cSQL & ", '" & cFirmKomm & "' as FIRMKOMM"
    cSQL = cSQL & " from REKOPF where SCHLUESSEL = '" & cReNr & "' and REDATUM = " & cDatum & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update DRU_REKO set Kommentar = '" & cKommentar & "'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into DRU_REPO "
    cSQL = cSQL & "Select * from REPOS where SCHLUESSEL = '" & cReNr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cNettoErm = Format(ermNettoERM, "#####0.00")
    cNettoErm = SwapStr(cNettoErm, ",", ".")
    
    cNettoVoll = Format(ermNettoVoll, "#####0.00")
    cNettoVoll = SwapStr(cNettoVoll, ",", ".")
    
    loeschNEW "NETTOS", gdBase
    
    cSQL = "Create Table NETTOS "
    cSQL = cSQL & "("
    cSQL = cSQL & " NettERM double"
    cSQL = cSQL & ", NettVol double"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert Into NETTOS (NETTERM , NETTVOL) values ( " & cNettoErm & "," & cNettoVoll & " ) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index SCHLUESSEL from DRU_REKO"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index SCHLUESSEL from DRU_REPO"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index SCHLUESSEL on DRU_REKO (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index SCHLUESSEL on DRU_REPO (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from DRU_REKO"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!PREISKZ) Then
            cRePreisKz = rsrs!PREISKZ
        Else
            cRePreisKz = "N"
        End If
    Else
        cRePreisKz = "N"
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from FIRMA"
    Set rsF = gdBase.OpenRecordset(cSQL)
    If Not rsF.EOF Then
        If Not IsNull(rsF!Steuernr) Then
            cSteuernr = rsF!Steuernr
        Else
            cSteuernr = "keine Angabe"
        End If
    End If
    rsF.Close
    
    cSQL = "Update DRU_REPO SET STEUERNR = '" & cSteuernr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from DRU_REPO"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    iMixMWSt = 0
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!MWST) Then
                cMWStKz = rsrs!MWST
            Else
                cMWStKz = ""
            End If
            Select Case cMWStKz
                Case Is = "V"
                    If Not iMixMWSt And 1 Then
                        iMixMWSt = iMixMWSt + 1
                    End If
                Case Is = "E"
                    If Not iMixMWSt And 2 Then
                        iMixMWSt = iMixMWSt + 2
                    End If
                Case Is = "O"
                    If Not iMixMWSt And 4 Then
                        iMixMWSt = iMixMWSt + 4
                    End If
                Case Else
                    'ignorieren
            End Select
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Select Case iMixMWSt
        Case 0, 1, 2, 4     'MWST-Kz nur V oder nur E oder nur O oder leer
            Select Case cRePreisKz
                Case Is = "N"
                    If Modul6.FindFile(gcDBPfad, "aWKL34l.rpt") Then
                        reportbildschirm "spez5", "aWKL34l"
                    Else
                        reportbildschirm "WKL005", "aWKL34b"
                    End If
                Case Is = "B"
                    If Modul6.FindFile(gcDBPfad, "aWKL34m.rpt") Then
                        reportbildschirm "spez5a", "aWKL34m"
                    Else
                        reportbildschirm "WKL005a", "aWKL34c"
                    End If
                Case Is = "O"
                    If Modul6.FindFile(gcDBPfad, "aWKL34n.rpt") Then
                        reportbildschirm "spez5b", "aWKL34n"
                    Else
                        reportbildschirm "WKL005b", "aWKL34d"
                    End If
                Case Is = "Z"
                    If Modul6.FindFile(gcDBPfad, "aWKL34o.rpt") Then
                        reportbildschirm "spez5c", "aWKL34o"
                    Else
                        reportbildschirm "WKL005c", "aWKL34e"
                    End If
            End Select
        
        Case 3, 5, 6, 7         'MWSt-Kz ist gemischt
            Select Case cRePreisKz
                Case Is = "N"
                    If Modul6.FindFile(gcDBPfad, "aWKL34p.rpt") Then
                        reportbildschirm "spez5d", "aWKL34p"
                    Else
                        reportbildschirm "WKL005d", "aWKL34f"
                    End If
                Case Is = "B"
                    If Modul6.FindFile(gcDBPfad, "aWKL34q.rpt") Then
                        reportbildschirm "spez5e", "aWKL34q"
                    Else
                        reportbildschirm "WKL005e", "aWKL34g"
                    End If
                Case Is = "O"
                    If Modul6.FindFile(gcDBPfad, "aWKL34r.rpt") Then
                        reportbildschirm "spez5f", "aWKL34r"
                    Else
                        reportbildschirm "WKL005f", "aWKL34h"
                    End If
                Case Is = "Z"
                    If Modul6.FindFile(gcDBPfad, "aWKL34s.rpt") Then
                        reportbildschirm "spez5g", "aWKL34s"
                    Else
                        reportbildschirm "WKL005g", "aWKL34i"
                    End If
            End Select
    
    End Select
    

Exit Sub
LOKAL_ERROR:
    If err.Number = 3295 Or err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeMahnung"
        Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub SchreibeDatenInRechnungWKL24(bReFuss As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim ctmp As String
    Dim cKdnr As String
    Dim cSQL  As String
    Dim cInto As String
    Dim cIntoOFPO As String
    Dim rsrs As Recordset
    Dim rsF As Recordset
    
    Dim lReDatum As Long
    Dim dSumme As Double
    Dim lFlag As Long
    
    'Zielfelder
    Dim cTitel      As String
    Dim cKdName1    As String
    Dim cKdName2    As String
    Dim cStrasse    As String
    Dim cPlz        As String
    Dim cOrt        As String
    Dim cReNr       As String
    Dim cReDatum    As String
    Dim cReText     As String
    Dim dReSumme    As Double
    Dim dPortoVerp  As Double
    Dim cKommentar  As String
    Dim cPreisKz    As String
    Dim cPfad       As String
    
    Dim cFirmName   As String
    Dim cFirmAdress As String
    Dim cFirmBank   As String
    Dim cFirmKomm   As String
    Dim cSteuernr   As String
    Dim cMWStKz     As String
    Dim iMixMWSt    As Integer
    
    Dim sErstellt_von       As String
    Dim sIhr_zeichen        As String
    Dim sZahlungs_Ziel      As String
    Dim sSpezialSatz        As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If bReFuss Then
        cSQL = "Select * from FIRMA"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Steuernr) Then
                cSteuernr = rsrs!Steuernr
            Else
                cSteuernr = ""
            End If
            If Not IsNull(rsrs!name) Then
                cFirmName = rsrs!name
            Else
                cFirmName = ""
            End If
            If Not IsNull(rsrs!strasse) Then
                cFirmAdress = rsrs!strasse
            Else
                cFirmAdress = ""
            End If
            If Not IsNull(rsrs!Plz) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & "   " & rsrs!Plz
                Else
                    cFirmAdress = rsrs!Plz
                End If
            End If
            If Not IsNull(rsrs!Ort) Then
                If cFirmAdress <> "" Then
                    cFirmAdress = cFirmAdress & " " & rsrs!Ort
                Else
                    cFirmAdress = rsrs!Ort
                End If
            End If
            If Not IsNull(rsrs!BankName) Then
                cFirmBank = rsrs!BankName
            Else
                cFirmBank = ""
            End If
            
            If Not IsNull(rsrs!BLZ) Then
                If rsrs!BLZ <> "" Then
                    cFirmBank = cFirmBank & "  BLZ " & rsrs!BLZ
                End If
            End If
            
            If Not IsNull(rsrs!Konto) Then
                If rsrs!Konto <> "" Then
                    cFirmBank = cFirmBank & "  Konto: " & rsrs!Konto
                End If
            End If
            
            If Not IsNull(rsrs!BIC) Then
                If rsrs!BIC <> "" Then
                    cFirmBank = cFirmBank & "  BIC " & rsrs!BIC
                End If
            End If
            
            If Not IsNull(rsrs!IBAN) Then
                If rsrs!IBAN <> "" Then
                    cFirmBank = cFirmBank & "  IBAN: " & rsrs!IBAN
                End If
            End If
            
            If Not IsNull(rsrs!Tel) Then
                cFirmKomm = "Tel.: " & rsrs!Tel
            Else
                cFirmKomm = ""
            End If
            If Not IsNull(rsrs!Fax) Then
                If cFirmKomm <> "" Then
                    cFirmKomm = cFirmKomm & "  Fax: " & rsrs!Fax
                Else
                    cFirmKomm = "Fax: " & rsrs!Fax
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        Else
            cFirmName = ""
            cFirmAdress = ""
            cFirmBank = ""
            cFirmKomm = ""
        End If
    
    End If
    
    
   

    'Druck-Datei löschen und neu erzeugen
    loeschNEW "DRU_REKO", gdBase
    
    cSQL = "Create Table DRU_REKO "
    cSQL = cSQL & "(SCHLUESSEL Text(15)"
    cSQL = cSQL & ", KUNDNR LONG"
    cSQL = cSQL & ", ANREDE Text(35)"
    cSQL = cSQL & ", KDNAME1 Text(71)"
    cSQL = cSQL & ", KDNAME2 Text(71)"
    cSQL = cSQL & ", STRASSE Text(35)"
    cSQL = cSQL & ", PLZ Text(7)"
    cSQL = cSQL & ", ORT Text(35)"
    cSQL = cSQL & ", RENR Text(15)"
    cSQL = cSQL & ", REDATUM Datetime"
    cSQL = cSQL & ", RETEXT Text(100)"
    cSQL = cSQL & ", RESUMME double"
    cSQL = cSQL & ", STATUS Text(1)"
    cSQL = cSQL & ", PORTOVERP double"
    cSQL = cSQL & ", KOMMENTAR memo"
    cSQL = cSQL & ", PREISKZ Text(1)"
    cSQL = cSQL & ", STEUERNR Text(35)"
    cSQL = cSQL & ", FIRMNAME Text(50)"
    cSQL = cSQL & ", FIRMADRESS Text(110)"
    cSQL = cSQL & ", FIRMBANK Text(200)"
    cSQL = cSQL & ", FIRMKOMM Text(100)"
    cSQL = cSQL & ", IHR_ZEICHEN Text(30)"
    cSQL = cSQL & ", ERSTELLT_VON Text(30)"
    cSQL = cSQL & ", ZAHLUNGS_ZIEL Text(250)"
    cSQL = cSQL & ", SPEZIALSATZ Text(250)"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
    '*** Kundendaten lesen ***
    cTitel = ""
    cKdName1 = ""
    cKdName2 = ""
    cStrasse = ""
    cPlz = ""
    cOrt = ""
    sErstellt_von = ""
    sIhr_zeichen = ""
    sZahlungs_Ziel = ""
    sSpezialSatz = ""
        
    ctmp = Label5(0).Caption
    If ctmp = "" Then
        ctmp = "0"
    End If
    ctmp = fnMoveComma2Point$(ctmp)
    dPortoVerp = Val(ctmp)
    cKommentar = Label5(1).Caption
    cKommentar = SwapStr(cKommentar, "'", " ")
    
    sErstellt_von = Label5(3).Caption
    sErstellt_von = SwapStr(sErstellt_von, "'", " ")
    
    sIhr_zeichen = Label5(4).Caption
    sIhr_zeichen = SwapStr(sIhr_zeichen, "'", " ")
    
    sZahlungs_Ziel = Label5(5).Caption
    
    sSpezialSatz = Label5(6).Caption
    sSpezialSatz = SwapStr(sSpezialSatz, "'", " ")
    
    cKdnr = Label2(0).Caption
    cTitel = Text2(6).Text  'Anrede
    
    cKdName1 = ""
    If Text2(7).Text <> "" Then 'Titel
        cKdName1 = cKdName1 & Text2(7).Text & " "
    End If
    
    If Text2(1).Text <> "" Then 'Vorname
        cKdName1 = cKdName1 & Text2(1).Text & " "
    End If
    
    cKdName1 = cKdName1 & Text2(0).Text  'Name
    
    cKdName2 = Text2(5).Text 'Firma
    
    cStrasse = Text2(2).Text
    cPlz = Text2(3).Text
    cOrt = Text2(4).Text
    
    lReDatum = Fix(Now)
    cReDatum = Trim$(Str$(lReDatum))
    
    '*** aktuelle Rechnungsnummer holen ***
    
    '***************************************************
    '* Ersetzt durch manuelle ReNu-Vergabe in frmWK24b
    '***************************************************
    cReNr = gcReNr
        
    cInto = "Insert into REKOPF "
    cIntoOFPO = "Insert into OFPO "
    cSQL = "( "
    cSQL = cSQL & "SCHLUESSEL"
    cSQL = cSQL & ", KUNDNR"
    cSQL = cSQL & ", ANREDE"
    cSQL = cSQL & ", KDNAME1"
    cSQL = cSQL & ", KDNAME2"
    cSQL = cSQL & ", STRASSE"
    cSQL = cSQL & ", PLZ"
    cSQL = cSQL & ", ORT"
    cSQL = cSQL & ", RENR"
    cSQL = cSQL & ", REDATUM"
    cSQL = cSQL & ", RETEXT"
    cSQL = cSQL & ", RESUMME"
    cSQL = cSQL & ", STATUS"
    cSQL = cSQL & ", PORTOVERP"
    cSQL = cSQL & ", KOMMENTAR"
    cSQL = cSQL & ", PREISKZ"
    cSQL = cSQL & ", IHR_ZEICHEN "
    cSQL = cSQL & ", ERSTELLT_VON "
    cSQL = cSQL & ", ZAHLUNGS_ZIEL "
    cSQL = cSQL & ", SPEZIALSATZ "
    
    cSQL = cSQL & ") values ("
    cSQL = cSQL & "'" & cReNr & "'"
    cSQL = cSQL & ", " & cKdnr & ""
    cSQL = cSQL & ", '" & cTitel & "'"
    cSQL = cSQL & ", '" & cKdName1 & "'"
    cSQL = cSQL & ", '" & cKdName2 & "'"
    cSQL = cSQL & ", '" & cStrasse & "'"
    cSQL = cSQL & ", '" & cPlz & "'"
    cSQL = cSQL & ", '" & cOrt & "'"
    cSQL = cSQL & ", '" & cReNr & "'"
    cSQL = cSQL & ", " & cReDatum & ""
    cSQL = cSQL & ", '" & cReText & "'"
    cSQL = cSQL & ", 0 "
    cSQL = cSQL & ", 'O' "
    cSQL = cSQL & ", " & Trim$(Str$(dPortoVerp)) & ""
    cSQL = cSQL & ", '" & cKommentar & "'"
    cSQL = cSQL & ", '" & gcRePreisKz & "'"
    cSQL = cSQL & ", '" & sIhr_zeichen & "'"
    cSQL = cSQL & ", '" & sErstellt_von & "'"
    cSQL = cSQL & ", '" & sZahlungs_Ziel & "'"
    cSQL = cSQL & ", '" & sSpezialSatz & "'"
    cSQL = cSQL & ") "
    
    gdBase.Execute cInto & cSQL, dbFailOnError
    
    
    
    
     
    
    
    
    
    
    
    'OFPO auch
    gdBase.Execute cIntoOFPO & cSQL, dbFailOnError
    
    cInto = "Insert into DRU_REKO "
    cSQL = "( "
    cSQL = cSQL & "SCHLUESSEL"
    cSQL = cSQL & ", KUNDNR"
    cSQL = cSQL & ", ANREDE"
    cSQL = cSQL & ", KDNAME1"
    cSQL = cSQL & ", KDNAME2"
    cSQL = cSQL & ", STRASSE"
    cSQL = cSQL & ", PLZ"
    cSQL = cSQL & ", ORT"
    cSQL = cSQL & ", RENR"
    cSQL = cSQL & ", REDATUM"
    cSQL = cSQL & ", RETEXT"
    cSQL = cSQL & ", RESUMME"
    cSQL = cSQL & ", STATUS"
    cSQL = cSQL & ", PORTOVERP"
    cSQL = cSQL & ", KOMMENTAR"
    cSQL = cSQL & ", PREISKZ"
    cSQL = cSQL & ", STEUERNR"
    cSQL = cSQL & ", FIRMNAME"
    cSQL = cSQL & ", FIRMADRESS"
    cSQL = cSQL & ", FIRMBANK"
    cSQL = cSQL & ", FIRMKOMM"
    
    cSQL = cSQL & ", IHR_ZEICHEN "
    cSQL = cSQL & ", ERSTELLT_VON "
    cSQL = cSQL & ", ZAHLUNGS_ZIEL "
    cSQL = cSQL & ", SPEZIALSATZ "
    cSQL = cSQL & ") values ("
    cSQL = cSQL & "'" & cReNr & "'"
    cSQL = cSQL & ", " & cKdnr & ""
    cSQL = cSQL & ", '" & cTitel & "'"
    cSQL = cSQL & ", '" & cKdName1 & "'"
    cSQL = cSQL & ", '" & cKdName2 & "'"
    cSQL = cSQL & ", '" & cStrasse & "'"
    cSQL = cSQL & ", '" & cPlz & "'"
    cSQL = cSQL & ", '" & cOrt & "'"
    cSQL = cSQL & ", '" & cReNr & "'"
    cSQL = cSQL & ", " & cReDatum & ""
    cSQL = cSQL & ", '" & cReText & "'"
    cSQL = cSQL & ", 0 "
    cSQL = cSQL & ", 'O' "
    cSQL = cSQL & ", " & Trim$(Str$(dPortoVerp)) & ""
    cSQL = cSQL & ", '" & cKommentar & "'"
    cSQL = cSQL & ", '" & gcRePreisKz & "'"
    cSQL = cSQL & ", '" & cSteuernr & "'"
    cSQL = cSQL & ", '" & cFirmName & "'"
    cSQL = cSQL & ", '" & cFirmAdress & "'"
    cSQL = cSQL & ", '" & cFirmBank & "'"
    cSQL = cSQL & ", '" & cFirmKomm & "'"
    
    cSQL = cSQL & ", '" & sIhr_zeichen & "'"
    cSQL = cSQL & ", '" & sErstellt_von & "'"
    cSQL = cSQL & ", '" & sZahlungs_Ziel & "'"
    cSQL = cSQL & ", '" & sSpezialSatz & "'"
    
    
    cSQL = cSQL & ") "
    gdBase.Execute cInto & cSQL, dbFailOnError
    


    MSFlexGrid2.Redraw = False
    dReSumme = 0
    lAnzRecords = MSFlexGrid2.Rows - 1

    For lAktRecord = 1 To lAnzRecords
        dSumme = 0
        MSFlexGrid2.Row = lAktRecord
        MSFlexGrid2.Col = 0
        ctmp = MSFlexGrid2.Text
        ctmp = Trim$(ctmp)
        If ctmp = "ausbuchen" Then
            MSFlexGrid2.Col = 9
            lFlag = Val(MSFlexGrid2.Text)
            MoveDaten2RechnungWKL24 lAktRecord, cReNr, dSumme, lFlag
            MSFlexGrid2.Row = lAktRecord
            MSFlexGrid2.Col = 0
            MSFlexGrid2.Text = "gelöscht"
            
            MSFlexGrid2.Col = 10
            MSFlexGrid2.Text = ""
            
            MSFlexGrid2.Col = 0
            
        End If
        dReSumme = dReSumme + dSumme
    Next lAktRecord
    MSFlexGrid2.Redraw = True
    
    MoveRePos2DruRePosWKL24 cReNr
    
    cSQL = "Update REKOPF set RESUMME = " & Trim$(Str$(dReSumme)) & " where SCHLUESSEL = '" & cReNr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    'ofpo auch
    cSQL = "Update OFPO set RESUMME = " & Trim$(Str$(dReSumme)) & " where SCHLUESSEL = '" & cReNr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    aktuali_newOFPO cReNr
    
    cSQL = "Update DRU_REKO set RESUMME = " & Trim$(Str$(dReSumme)) & " where SCHLUESSEL = '" & cReNr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index SCHLUESSEL on DRU_REKO (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Create Index SCHLUESSEL on DRU_REPO (SCHLUESSEL)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from FIRMA"
    Set rsF = gdBase.OpenRecordset(cSQL)
    If Not rsF.EOF Then
        If Not IsNull(rsF!Steuernr) Then
            cSteuernr = rsF!Steuernr
        Else
            cSteuernr = "keine Angabe"
        End If
    End If
    rsF.Close
        
    cSQL = "Update DRU_REPO SET STEUERNR = '" & cSteuernr & "' "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    
    
    
    

    
    cSQL = "Select * from DRU_REPO"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    iMixMWSt = 0
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!MWST) Then
                cMWStKz = rsrs!MWST
            Else
                cMWStKz = ""
            End If
            Select Case cMWStKz
                Case Is = "V"
                    If Not iMixMWSt And 1 Then
                        iMixMWSt = iMixMWSt + 1
                    End If
                Case Is = "E"
                    If Not iMixMWSt And 2 Then
                        iMixMWSt = iMixMWSt + 2
                    End If
                Case Is = "O"
                    If Not iMixMWSt And 4 Then
                        iMixMWSt = iMixMWSt + 4
                    End If
                Case Else
                    'ignorieren
            End Select
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
    Select Case gcRePreisKz
        Case Is = "N"
            If Modul6.FindFile(gcDBPfad, "aWKL24l.rpt") Then
                reportbildschirm "spez5", "aWKL24l"
            Else
                reportbildschirm "WKL005", "aWKL24b"
            End If
        Case Is = "B"
            If Modul6.FindFile(gcDBPfad, "aWKL24m.rpt") Then
            
            
                
    
                'If Modul6.FindFile(gcDBPfad, "aWKL24m.rpt") And UCase(gsStammFTPUSER) = "HICKMANN" Then
                     
                     'bau mal eine Tabelle für die regulären KVK und ArtRab in Proz
                    
                    cSQL = "Alter Table DRU_REPO add ARTRAB Double "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Alter Table DRU_REPO add KVKPR1 Double "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Update DRU_REPO inner join Artikel on DRU_REPO.Artnr = Artikel.Artnr "
                    cSQL = cSQL & " Set DRU_REPO.KVKPR1 = Artikel.KVKPR1 "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Update DRU_REPO SET Artrab = 100 - (EPREIS*100/kvkpr1) where kvkpr1 <> 0  "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    cSQL = "Update DRU_REPO SET Artrab = Artrab * (-1)  "
                    gdBase.Execute cSQL, dbFailOnError
                    
                    'ARTRAB , KVKPR1 formatieren
                    cSQL = "UPDATE DRU_REPO SET ARTRAB=FORMAT(ARTRAB,'0.00') , KVKPR1=FORMAT(KVKPR1,'0.00')"
                    gdBase.Execute cSQL, dbFailOnError
           
           
                    'Ende, bau mal eine Tabelle
                    
                'End If
                
                
            
            
            
            
            
            
            
                reportbildschirm "spez5a", "aWKL24m"
            Else
                reportbildschirm "WKL005a", "aWKL24c"
            End If
        Case Is = "O"
            If Modul6.FindFile(gcDBPfad, "aWKL24n.rpt") Then
                reportbildschirm "spez5b", "aWKL24n"
            Else
                reportbildschirm "WKL005b", "aWKL24d"
            End If
        Case Is = "Z"
            If Modul6.FindFile(gcDBPfad, "aWKL24o.rpt") Then
                reportbildschirm "spez5c", "aWKL24o"
            Else
                reportbildschirm "WKL005c", "aWKL24e"
            End If
        
    End Select
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3372 Or err.Number = 53 Or err.Number = 3376 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeDatenInRechnungWKL24"
        Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Sub
Private Sub ZeigeVorhandeneRechnungenWKL24()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lAnzRecords As Long
    Dim lAktRecord As Long
    Dim ctmp As String
    Dim dWert As Double
    
    MSFlexGrid3.Visible = False
    
    MSFlexGrid3.Rows = 1
    
    MSFlexGrid3.Rows = 2
    
    MSFlexGrid3.Row = 0
    MSFlexGrid3.Col = 0
    MSFlexGrid3.Text = "Re-Datum"
    MSFlexGrid3.ColWidth(0) = 1000
    MSFlexGrid3.Col = 1
    MSFlexGrid3.Text = "Re-Nr"
    MSFlexGrid3.ColWidth(1) = 1000
    MSFlexGrid3.Col = 2
    MSFlexGrid3.Text = "KdNr"
    MSFlexGrid3.ColWidth(2) = 700
    MSFlexGrid3.Col = 3
    MSFlexGrid3.Text = "Kundenname"
    MSFlexGrid3.ColWidth(3) = 2000
    MSFlexGrid3.Col = 4
    MSFlexGrid3.Text = "Straße"
    MSFlexGrid3.ColWidth(4) = 2000
    MSFlexGrid3.Col = 5
    MSFlexGrid3.Text = "PLZ"
    MSFlexGrid3.ColWidth(5) = 800
    MSFlexGrid3.Col = 6
    MSFlexGrid3.Text = "Ort"
    MSFlexGrid3.ColWidth(6) = 2000
    MSFlexGrid3.Col = 7
    MSFlexGrid3.Text = "Rechnungssumme"
    MSFlexGrid3.ColWidth(7) = 1500
    MSFlexGrid3.Col = 8
    MSFlexGrid3.Text = "Status"
    MSFlexGrid3.ColWidth(8) = 1000
    
    
    cSQL = "Select * from REKOPF order by REDATUM desc, RENR desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzRecords = rsrs.RecordCount
        MSFlexGrid3.Rows = lAnzRecords + 1
        rsrs.MoveFirst
        lAktRecord = 0
        Do While Not rsrs.EOF
            lAktRecord = lAktRecord + 1
            MSFlexGrid3.Row = lAktRecord
            
            If Not IsNull(rsrs!REDATUM) Then
                ctmp = rsrs!REDATUM
            Else
                ctmp = ""
            End If
            MSFlexGrid3.Col = 0
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!RENR) Then
                ctmp = rsrs!RENR
            Else
                ctmp = ""
            End If
            MSFlexGrid3.Col = 1
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!Kundnr) Then
                ctmp = rsrs!Kundnr
            Else
                ctmp = ""
            End If
            MSFlexGrid3.Col = 2
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!KDNAME1) Then
                ctmp = rsrs!KDNAME1
            Else
                ctmp = ""
            End If
            MSFlexGrid3.Col = 3
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!strasse) Then
                ctmp = rsrs!strasse
            Else
                ctmp = ""
            End If
            MSFlexGrid3.Col = 4
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!Plz) Then
                ctmp = rsrs!Plz
            Else
                ctmp = ""
            End If
            MSFlexGrid3.Col = 5
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!Ort) Then
                ctmp = rsrs!Ort
            Else
                ctmp = ""
            End If
            MSFlexGrid3.Col = 6
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!RESUMME) Then
                dWert = rsrs!RESUMME
            Else
                dWert = 0
            End If
            ctmp = Format$(dWert, "###,##0.00") & " " & gcWaehrung
            MSFlexGrid3.Col = 7
            MSFlexGrid3.Text = ctmp
            
            If Not IsNull(rsrs!Status) Then
                ctmp = rsrs!Status
            Else
                ctmp = "O"
            End If
            MSFlexGrid3.Col = 8
            MSFlexGrid3.Text = ctmp
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Frame3.Visible = True
    Frame1.Visible = False
    
    TabellenbreiteanpassenNH MSFlexGrid3, 1.25 * gdTabfak
    
    MSFlexGrid3.Visible = True
    
Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeVorhandeneRechnungenWKL24"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "REDATUM"
                SpaltennummerReDatum = i
            Case Is = "ZAHLUNGSZIEL"
                SpaltennummerZahlZiel = i
            Case Is = "KUNDNR"
                SpaltennummerKdNr = i
            Case Is = "RENR"
                SpaltennummerReNr = i
            Case Is = "STATUSBEZ"
                SpaltennummerStatusBez = i
            Case Is = "ZAHLUNGSINFO"
                SpaltennummerBezahlInfo = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Zeige_OffenePosten()
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    
    Screen.MousePointer = 11

    Tabcheck "OFPO"
    
    FormatGridOverTablay "OFPO"

    With MSFlexGrid4
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 2
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
        Next j
    
        FuellenMSFlex_OFPO

        ermittlespalten
        
        .Redraw = False
        
        TabellenbreiteanpassenNH MSFlexGrid4, 1.25 * gdTabfak
        
        .Visible = True
        .Redraw = True
        .Row = 1
    End With
    
    Frame4.Visible = True
    Frame1.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zeige_OffenePosten"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub setzeZahlZiel(iZahlZiel As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    
    With MSFlexGrid4
        
        For i = 0 To .Rows - 1
            If .TextMatrix(i, SpaltennummerReDatum) <> "" And .TextMatrix(i, SpaltennummerReDatum) <> "Re-Datum" Then
                .TextMatrix(i, SpaltennummerZahlZiel) = Format(DateValue(.TextMatrix(i, SpaltennummerReDatum)) + iZahlZiel, "DD.MM.YY")
            End If
            
'            If DateValue(.TextMatrix(i, SpaltennummerReDatum)) < DateValue(Now) + iZahlZiel Then
'
'
'            End If
        Next i
        
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "setzeZahlZiel"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMSFlex_OFPO()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim cSQL        As String
    Dim corder      As String
    Dim cwhere      As String
    
    If Option1(2).value Then
         corder = " order by ZAHLUNGSZIEL desc , val(RENR) desc"
    ElseIf Option1(0).value Then
         corder = " order by val(RENR) "
    ElseIf Option1(1).value Then
        corder = " order by KUNDNR "
    End If

    If Option1(4).value Then
        cwhere = " "
    ElseIf Option1(5).value Then
        cwhere = " where Statusbez = 'nicht bezahlt' "
    ElseIf Option1(7).value Then
        cwhere = " where Statusbez = 'bezahlt' "
    End If

    cSQL = "Select * from OFPO " & cwhere & corder
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    With MSFlexGrid4
        .Redraw = False
        lrow = 1
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                lrow = lrow + 1
                .Rows = lrow + 1
                .Col = 0
                
                For i = 0 To byAnzahlSpalten - 1
                    .Row = 0
                    .Col = i
                    
                    If sSpaltenname(i) = .Text Then
                        Select Case sSpaltenname(i)
                                
                            Case Is = "Re-Summe"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = "0"
                                End If
                                .Row = lrow
                                .Text = Format$(sWert, "#######0.00")
                                
                            Case Is = "überschritten"
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                
                                If Left(sWert, 1) = "-" Or Left(sWert, 1) = "0" Then
                                    .Text = ""
                                    .CellForeColor = vbBlack
                                Else
                                    .Text = sWert
                                    .CellForeColor = vbRed
                                End If
    
                            Case Else
                                If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                    sWert = rsrs(sSpaltenbez(i))
                                Else
                                    sWert = ""
                                End If
                                .Row = lrow
                                .Text = sWert
                        End Select
                        
                
                        If Len(.TextMatrix(lrow, i)) * 80 > aBreite(i) Then
                            aBreite(i) = Len(.TextMatrix(lrow, i)) * 80
                        End If
                        
                    End If
                Next i
                rsrs.MoveNext
            Loop
        End If
        
        For i = 0 To byAnzahlSpalten - 1
            .Col = i
            .ColWidth(i) = aBreite(i) * 1.8
        Next i
            
        
        rsrs.Close: Set rsrs = Nothing
        
        If byAnzahlSpalten < 2 Then
        
        Else
            .FixedCols = 1
        End If
        
        .RowHeight(1) = 0
        lrow = lrow - 1
        .Redraw = True
        .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMSFlex_OFPO"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
    Fehlermeldung1
End Sub
Private Sub Check1_Click()
    On Error GoTo LOKAL_ERROR
    
    If Check1.value = vbChecked Then
        Check1.Caption = "alle zurücksetzen"
        flex "zurücksetzen"
    Else
        Check1.Caption = "alle markieren"
        flex "ausbuchen"
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub flex(krit As String)
On Error GoTo LOKAL_ERROR
    
    Dim lcount  As Long
    
    MSFlexGrid2.Redraw = False
    For lcount = 1 To MSFlexGrid2.Rows - 1
        MSFlexGrid2.Col = 0
        MSFlexGrid2.Row = lcount
        
        If Left(MSFlexGrid2.Text, 8) <> "gelöscht" Then

        
            Select Case krit
                Case "zurücksetzen"
                
                    MSFlexGrid2.Text = "ausbuchen"
                    MSFlexGrid2.Col = 10
                    MSFlexGrid2.Text = lcount
    
                Case "ausbuchen"
                    MSFlexGrid2.Text = "offen"
                    MSFlexGrid2.Col = 10
                    MSFlexGrid2.Text = ""
                
            End Select
        End If
        
    Next lcount
    
    MSFlexGrid2.Redraw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "flex"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim sSQL As String

    Select Case index
        Case 0   '//Details
            If Label0.Caption = "-1" Then
                MsgBox "Bitte eine Zeile in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
                MSFlexGrid1.SetFocus
            Else
                LeseKreditDetailsWKL24
                
                If NewTableSuchenDBKombi("E24", gdApp) Then
                    voreinstellungladen24
                End If
                
                Command2(3).Caption = "Zusatz Rechnung 1"
                Command2(3).ForeColor = glButtonForecolor
                Label5(2).Caption = ""
            End If
        Case 1  '//Schließen
            Unload frmWKL24
        Case 2   '//alte Rechnungen anlisten
            ZeigeVorhandeneRechnungenWKL24
        Case 3  '//Liste drucken
            DruckeOffeneKrediteWKL24
        Case 4  'offene Postenliste
            If NewTableSuchenDBKombi("E24O", gdApp) Then
                voreinstellungladen24o
            End If
            
            aktualisiere_ZahlZiel
            
            Zeige_OffenePosten
    End Select
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim dSumme As Double
    Dim dWert As Double
    Dim lcount As Long
    Dim cFeld As String
    
    Screen.MousePointer = 11
    
    Select Case index
        Case Is = 0     'Ausbuchen   '//Kunden bezahlt
            dSumme = 0
            dWert = 0
            
            For lcount = 1 To MSFlexGrid2.Rows - 1
                MSFlexGrid2.Row = lcount
                MSFlexGrid2.Col = 0
                cFeld = MSFlexGrid2.Text
                cFeld = Trim$(cFeld)
                If cFeld = "ausbuchen" Then
                    MSFlexGrid2.Col = 6
                    cFeld = MSFlexGrid2.Text
                    cFeld = Trim$(cFeld)
                    cFeld = fnMoveComma2Point$(cFeld)
                    dWert = Val(cFeld)
                    dSumme = dSumme + dWert
                End If
            Next lcount
            
            If dSumme = 0 Then
                Screen.MousePointer = 0
                MsgBox "Bitte einen Eintrag zum Ausbuchen auswählen!", vbCritical, "Winkiss Hinweis:"
                MSFlexGrid2.SetFocus
                Exit Sub
            End If
            
            If Check3.value = vbChecked Then
                AusbuchenKreditVerkaufWKL24ohneBon
            Else
            
'                HoleNeueBonNrWKL20
                HoleNeueBonNrWKL20_NEU
                AusbuchenKreditVerkaufWKL24
            End If
            

            
            setzedrucker gcListenDrucker
            
        Case Is = 1     'Rechnung schreiben
        
            voreinstellungspeichern24
        
            Command2(0).Enabled = False
            Check3.Visible = False
        
            iRet = fnPruefeStatusOffeneKrediteWKL24%()
            If iRet = 0 Then
                frmWK24b.Show 1
                If Trim$(gcReNr) <> "" Then
                    If Check4.value = vbChecked Then
                        SchreibeDatenInRechnungWKL24 True
                    Else
                        SchreibeDatenInRechnungWKL24 False
                    End If
                Else
'                    MsgBox "Keine Rechnungsnummer vorgegeben!", vbCritical, "Winkiss Hinweis:"
                    Command2(index).SetFocus
                End If
            End If
        
        Case Is = 2     'Zurück
        
            If Check2.value = vbChecked Then
                LeseOffeneKrediteWKL24 0
            Else
                LeseOffeneKrediteWKL24 61
            End If
            Frame1.Visible = True
            Frame2.Visible = False
            
        Case Is = 3     'Zusätze Rechnung
            Command2(3).ForeColor = glButtonForecolor
            frmWK24a.Show 1
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub aktualisiere_ZahlZiel()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    sSQL = "Update OFPO set ZAHLUNGSZIEL = Datevalue(REDATUM) + " & Val(Text1.Text) & " "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update OFPO set UEBERSCHREITUNG = cstr(Datevalue(now)- Datevalue(ZAHLUNGSZIEL)) + ' Tage überschritten'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update OFPO set UEBERSCHREITUNG = ''"
    sSQL = sSQL & " where Statusbez = 'bezahlt'"
    gdBase.Execute sSQL, dbFailOnError
    
    Option1(5).Caption = "nicht bezahlte" & " (" & Format(ermResum_nichtbezahlt, "###,##0.00") & ")"
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "aktualisiere_ZahlZiel"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermResum_nichtbezahlt() As Double
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermResum_nichtbezahlt = 0
    sSQL = "select sum(resumme) as maxi from OFPO where statusbez = 'nicht bezahlt'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermResum_nichtbezahlt = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
         
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermsumbest"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub Command3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iRet    As Integer
    Dim sSQL    As String
    Dim cVon    As String
    Dim cBis    As String
    Dim lVon    As Long
    Dim lBis    As Long
    
    Screen.MousePointer = 11
    
    Select Case index
        Case Is = 0
            If Label4.Caption = "-1" Then
                MsgBox "Bitte eine Rechnung auswählen!", vbCritical, "STOP!"
            Else
                If Check5.value = vbChecked Then
                    SchreibeAltRechnungWKL24 True
                Else
                    SchreibeAltRechnungWKL24 False
                End If
            End If
        Case Is = 1
            Label4.Caption = "-1"
            Frame1.Visible = True
            Frame3.Visible = False
         Case Is = 10
            voreinstellungspeichern24o
            Label4.Caption = "-1"
            Frame1.Visible = True
            Frame4.Visible = False
        Case Is = 2
            If Label4.Caption = "-1" Then
                MsgBox "Bitte eine Rechnung auswählen!", vbCritical, "STOP!"
            Else
                RechnungRueckgaengigWKL24
            End If
        Case Is = 3
            If Label4.Caption = "-1" Then
                MsgBox "Bitte eine Rechnung auswählen!", vbCritical, "STOP!"
            Else
                RechnungLoeschenWKL24
            End If
        Case Is = 4
            DruckeAlteKrediteWKL24
        Case Is = 5
            If Label4.Caption = "-1" Then
                MsgBox "Bitte eine Rechnung auswählen!", vbCritical, "STOP!"
            Else
                frmWKL132.Show 1
            End If
        Case 6
        
            cVon = Text21(1).Text
            cBis = Text21(2).Text
            
            If cVon <> "" Then
                lVon = DateValue(cVon)
                cVon = Trim$(Str$(lVon))
            End If
            
            If cBis <> "" Then
                lBis = DateValue(cBis)
                cBis = Trim$(Str$(lBis))
            End If
            
            Exportiere_Vorh_Rechnungen cVon, cBis
        Case 12
        
            aktualisiere_ZahlZiel
            
            Zeige_OffenePosten
    End Select
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ExportCSV()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    Dim i               As Integer

   
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(8)
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sSQL = " Select "
    sSQL = sSQL & " SCHLUESSEL  "
    sSQL = sSQL & ", KUNDNR  "
    sSQL = sSQL & ", KDNAME1  "
    sSQL = sSQL & ", KDNAME2  "
    sSQL = sSQL & ", STRASSE  "
    sSQL = sSQL & ", PLZ  "
    sSQL = sSQL & ", ORT  "
    sSQL = sSQL & ", RENR  "
    sSQL = sSQL & ", REDATUM  "
    sSQL = sSQL & ", RESUMME  "
    sSQL = sSQL & " from RE_EXPORT "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        sAusgabedatname = "Rechnungen.csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = "SCHLUESSEL;KUNDNR;KDNAME1;KDNAME2;STRASSE;PLZ;ORT;RENR;REDATUM;RESUMME" & Chr$(13) & Chr$(10)

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 9
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > 0 Then
                        If i = 8 Then
                            cSatz = cSatz & ";" & Format(rsrs.Fields(i), "DD.MM.YY")
                        
                        ElseIf i = 9 Then
                        
                            If rsrs.Fields(i) = 0 Then
                                cSatz = cSatz & ";"
                            Else
                                cSatz = cSatz & ";" & Format(rsrs.Fields(i), "###,##0.00")
                            End If
                        Else
                            cSatz = cSatz & ";" & rsrs.Fields(i)
                        End If
                    Else
                        cSatz = rsrs.Fields(i)
                    End If
                Else
                    If i > 0 Then
                        cSatz = cSatz & ";"
                    Else
                        cSatz = ""
                    End If
                End If
            Next i
        
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            rsrs.MoveNext
        Loop
        
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("RE_EXPORT", gdBase) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(8)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(8)
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
        Fehler.gsFunktion = "ExportCSV"
        Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Exportiere_Vorh_Rechnungen(cVon As String, cBis As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim bAnd As Boolean
    
    loeschNEW "RE_EXPORT", gdBase
    CreateTableT2 "RE_EXPORT", gdBase
    
    
    
    
    bAnd = False
    

    cSQL = "Insert into RE_EXPORT Select "
    cSQL = cSQL & " SCHLUESSEL "
    cSQL = cSQL & ", KUNDNR "
    cSQL = cSQL & ", KDNAME1 "
    cSQL = cSQL & ", KDNAME2 "
    cSQL = cSQL & ", STRASSE "
    cSQL = cSQL & ", PLZ "
    cSQL = cSQL & ", ORT "
    cSQL = cSQL & ", RENR "
    cSQL = cSQL & ", REDATUM "
    cSQL = cSQL & ", RESUMME "
    cSQL = cSQL & " from OFPO "
    
    If cVon <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & "  REDATUM >= " & cVon
        bAnd = True
    End If
    
    If cBis <> "" Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & "  REDATUM <= " & cBis
        bAnd = True
    End If
    
    
        
    If Option1(4).value Then
        
    ElseIf Option1(5).value Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " Statusbez = 'nicht bezahlt' "
    ElseIf Option1(7).value Then
        If bAnd Then
            cSQL = cSQL & " and "
        Else
            cSQL = cSQL & " where "
        End If
        cSQL = cSQL & " Statusbez = 'bezahlt' "
    End If
    
    
    cSQL = cSQL & " order by REDATUM desc, RENR desc"
    gdBase.Execute cSQL, dbFailOnError
    
    ExportCSV
    

Exit Sub

LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Exportiere_Vorh_Rechnungen"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub

Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnz As Long
    Dim lcount As Long
    
    Screen.MousePointer = 11
    
    PositionierenWKL24
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    
    
    If NewTableSuchenDBKombi("OFPO", gdBase) = False Then
         CreateTableT2 "OFPO", gdBase
         
    Else
        If SpalteInTabellegefundenNEW("OFPO", "IHR_ZEICHEN", gdBase) = False Then
            SpalteAnfuegenNEW "OFPO", "IHR_ZEICHEN", "Text(30)", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("OFPO", "ERSTELLT_VON", gdBase) = False Then
            SpalteAnfuegenNEW "OFPO", "ERSTELLT_VON", "Text(30)", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("OFPO", "ZAHLUNGS_ZIEL", gdBase) = False Then
            SpalteAnfuegenNEW "OFPO", "ZAHLUNGS_ZIEL", "Text(250)", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("OFPO", "SPEZIALSATZ", gdBase) = False Then
            SpalteAnfuegenNEW "OFPO", "SPEZIALSATZ", "Text(250)", gdBase
        End If
    End If
    
    If NewTableSuchenDBKombi("REKOPF", gdBase) = False Then
         CreateTableT2 "REKOPF", gdBase
    Else
        If SpalteInTabellegefundenNEW("REKOPF", "IHR_ZEICHEN", gdBase) = False Then
            SpalteAnfuegenNEW "REKOPF", "IHR_ZEICHEN", "Text(30)", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("REKOPF", "ERSTELLT_VON", gdBase) = False Then
            SpalteAnfuegenNEW "REKOPF", "ERSTELLT_VON", "Text(30)", gdBase
        End If
        
        If SpalteInTabellegefundenNEW("REKOPF", "ZAHLUNGS_ZIEL", gdBase) = False Then
            SpalteAnfuegenNEW "REKOPF", "ZAHLUNGS_ZIEL", "Text(250)", gdBase
        End If
        
        
        If SpalteInTabellegefundenNEW("REKOPF", "SPEZIALSATZ", gdBase) = False Then
            SpalteAnfuegenNEW "REKOPF", "SPEZIALSATZ", "Text(250)", gdBase
        End If
    End If
    
    LeseOffeneKrediteWKL24 61
    
    Label0.Caption = "1"
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern24()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    Dim bo0 As Integer
    Dim bo1 As Integer

    loeschNEW "E24", gdApp
    CreateTableT2 "E24", gdApp

    If Check3.value = vbChecked Then
        bo0 = 0
    Else
        bo0 = -1
    End If
    
    If Check4.value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If
    
    sSQL = "Insert into E24 ( bo0,bo1) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & ")"
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern24"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern24o()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    Dim iZhlZiel As Integer
    Dim bo0 As Integer
    Dim bo1 As Integer
    Dim bo2 As Integer
    Dim bo3 As Integer
    Dim bo4 As Integer
    Dim bo5 As Integer

    loeschNEW "E24O", gdApp
    CreateTableT2 "E24O", gdApp

    If Text1.Text <> "" Then
        iZhlZiel = Val(Text1.Text)
    End If
    
    bo0 = Option1(4).value
    bo1 = Option1(5).value
    bo2 = Option1(7).value
    bo3 = Option1(2).value
    bo4 = Option1(0).value
    bo5 = Option1(1).value
    

    
    sSQL = "Insert into E24O (bo0,bo1,bo2,bo3,bo4,bo5,ZhlZiel) "
    sSQL = sSQL & " values (" & bo0 & "," & bo1 & "," & bo2 & "," & bo3 & "," & bo4
    sSQL = sSQL & " ," & bo5 & "," & iZhlZiel & ")"
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern24o"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen24()
    On Error GoTo LOKAL_ERROR

    Dim rs As Recordset

    Set rs = gdApp.OpenRecordset("E24")
    If Not rs.EOF Then
        If rs!bo0 = True Then
            Check3.value = vbUnchecked
        Else
            Check3.value = vbChecked
        End If

        If rs!bo1 = True Then
            Check4.value = vbUnchecked
        Else
            Check4.value = vbChecked
        End If
    End If
    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen24"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen24o()
    On Error GoTo LOKAL_ERROR

    Dim rs As Recordset

    Set rs = gdApp.OpenRecordset("E24O")
    If Not rs.EOF Then
        If Not IsNull(rs!ZHLZIEL) Then
            Text1.Text = Val(rs!ZHLZIEL)
        Else
            Text1.Text = 14
        End If
        
        Option1(4).value = rs!bo0
        Option1(5).value = rs!bo1
        Option1(7).value = rs!bo2
        
        Option1(2).value = rs!bo3
        Option1(0).value = rs!bo4
        Option1(1).value = rs!bo5
        
    End If
    rs.Close: Set rs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen24o"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

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
    Fehler.gsFehlertext = "Beim Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub TabellenbreiteanpassenNH(gridx As MSFlexGrid, siEigFak As Single)
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
    Fehler.gsFunktion = "TabellenbreiteanpassenNH"
    Fehler.gsFehlertext = "Beim Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid1_Click()
    On Error GoTo LOKAL_ERROR
    
    Label0.Caption = MSFlexGrid1.Row
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_DblClick()
On Error GoTo LOKAL_ERROR
    
    If MSFlexGrid2.Row = 1 Then
        sortierenGrid MSFlexGrid2
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSFlexGrid2.Row
    lcol = MSFlexGrid2.Col

    Select Case lcol
        Case Is = 3, 5, 10
        If iKeypress = 0 And KeyCode <> vbKeyBack Then
            MSFlexGrid2.Row = lrow
            MSFlexGrid2.Col = lcol
            MSFlexGrid2.Text = ""
        End If
        iKeypress = iKeypress + 1
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcol As Long
    
    MSFlexGrid2.Redraw = False
    
    lcol = MSFlexGrid2.Col
    
    If lcol = 1 Then
        Label3.Caption = MSFlexGrid2.Row
        MSFlexGrid2.Col = 0
        ctmp = MSFlexGrid2.Text
        ctmp = Trim$(ctmp)
        
        Select Case ctmp
            Case Is = "offen"
                MSFlexGrid2.Text = "ausbuchen"
                If bUnterschiedlicheRechnunganschriften = True Then
                    spezRechnungsanschriftzeigen MSFlexGrid2.Row
                End If
                
                SetzenextReihenfolgennumber CLng(MSFlexGrid2.Row)
                
            Case Is = "ausbuchen"
                MSFlexGrid2.Text = "offen"
                
                MSFlexGrid2.Col = 10
                MSFlexGrid2.Text = ""
                
                MSFlexGrid2.Col = 1
            Case Else
                'nix tun
        End Select
    Else
    
    End If
    MSFlexGrid2.Redraw = True
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SetzenextReihenfolgennumber(lZeile As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount  As Long
    Dim lMax    As Long
    Dim lri     As Long
    
    lMax = 0
    
    'grid abklappern lmax ermitteln
    
    For lcount = 0 To MSFlexGrid2.Rows - 1
        MSFlexGrid2.Col = 10
        MSFlexGrid2.Row = lcount
        lri = CLng(Val(MSFlexGrid2.Text))
        If lri > lMax Then
            lMax = lri
        Else
        
        End If
    Next lcount
    
    
    
    'lmax in lzeile setzen
    
    MSFlexGrid2.Col = 10
    MSFlexGrid2.Row = lZeile
    MSFlexGrid2.Text = lMax + 1
    lMax = 0
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SetzenextReihenfolgennumber"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcol        As Long
    Dim lrow        As Long
    Dim cFeld       As String
    Dim cValid      As String
    Dim cZeichen    As String
    
    lcol = MSFlexGrid2.Col
    lrow = MSFlexGrid2.Row
    cFeld = MSFlexGrid2.Text
    
    
    
    
   
    
    Select Case lcol
        Case Is = 3
            Select Case KeyAscii
                Case Is = 8
                    If Len(cFeld) > 0 Then
                        cFeld = Left(cFeld, Len(cFeld) - 1)
                    End If
                Case Is = 13
                    cFeld = cFeld
                
                Case Else
                    cFeld = cFeld & Chr$(KeyAscii)
            End Select
            
            MSFlexGrid2.Text = cFeld
            
        Case Is = 5
        
            cZeichen = Chr$(KeyAscii)
        
        
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            


            If KeyAscii <> 0 Then
            MSFlexGrid2.Row = lrow
            MSFlexGrid2.Col = lcol
            cValid = MSFlexGrid2.Text
            If InStr(cValid, ",") > 0 And cZeichen = "," Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                If KeyAscii <> 8 Then
                    cValid = cValid & Chr$(KeyAscii)
                Else
                    If Len(cValid) > 0 Then
                        cValid = Left(cValid, Len(cValid) - 1)
                    End If
                End If
                
                MSFlexGrid2.Text = cValid
                
            End If
        
    End If
        
        Case Is = 10
        
            cZeichen = Chr$(KeyAscii)
        
        
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            


            If KeyAscii <> 0 Then
            MSFlexGrid2.Row = lrow
            MSFlexGrid2.Col = lcol
            cValid = MSFlexGrid2.Text
            If InStr(cValid, ",") > 0 And cZeichen = "," Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                If KeyAscii <> 8 Then
                    cValid = cValid & Chr$(KeyAscii)
                Else
                    If Len(cValid) > 0 Then
                        cValid = Left(cValid, Len(cValid) - 1)
                    End If
                End If
                
                MSFlexGrid2.Text = cValid
                
            End If
        
    End If
        
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub MSFlexGrid2_LeaveCell()
    On Error GoTo LOKAL_ERROR

    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSFlexGrid2.Row
    lcol = MSFlexGrid2.Col
    
    Select Case lcol
        Case Is = 3, 5, 10
            Select Case KeyCode
                Case Is = 46    'Del
                    MSFlexGrid2.Text = ""
                    
            End Select
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub spezRechnungsanschriftzeigen(lrow As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim iAdressID   As Integer
    Dim sKunr       As String
    Dim sArtnr      As String
    Dim lDatum      As Long
    Dim rsrs        As Recordset
    
    sKunr = Label2(0).Caption
    
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Row = lrow
    sArtnr = MSFlexGrid2.Text
    MSFlexGrid2.Col = 1
    lDatum = DateValue(MSFlexGrid2.Text)
    
    sSQL = "Select * from Kredit where KUNDNR = " & sKunr
    sSQL = sSQL & " and artnr = " & sArtnr
    sSQL = sSQL & " and adate = " & lDatum
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        iAdressID = IIf(IsNull(rsrs!AdressID), 0, rsrs!AdressID)
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If iAdressID > 0 Then
        sSQL = "Select * from zadress where Adressid = " & iAdressID
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            Text2(5).Text = IIf(IsNull(rsrs!firma), "", rsrs!firma)
            Text2(6).Text = IIf(IsNull(rsrs!anrede), "", rsrs!anrede)
            Text2(0).Text = IIf(IsNull(rsrs!name), "", rsrs!name)
            Text2(1).Text = IIf(IsNull(rsrs!vorname), "", rsrs!vorname)
            Text2(2).Text = IIf(IsNull(rsrs!strasse), "", rsrs!strasse)
            Text2(3).Text = IIf(IsNull(rsrs!Plz), "", rsrs!Plz)
            Text2(4).Text = IIf(IsNull(rsrs!Ort), "", rsrs!Ort)
            Text2(7).Text = IIf(IsNull(rsrs!titel), "", rsrs!titel)
                
        End If
        rsrs.Close: Set rsrs = Nothing
    
    Else
    
        Text2(0).Text = Label2(1).Caption
        Text2(1).Text = Label2(6).Caption
        Text2(2).Text = Label2(2).Caption
        Text2(3).Text = Label2(3).Caption
        Text2(4).Text = Label2(4).Caption
        Text2(5).Text = cFirma
        Text2(6).Text = cAnrede
        Text2(7).Text = cTitel
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "spezRechnungsanschriftzeigen"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid3_Click()
    On Error GoTo LOKAL_ERROR
    
    Label4.Caption = MSFlexGrid3.Row
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid4_Click()
On Error GoTo LOKAL_ERROR

    Dim sReNr As String

    If MSFlexGrid4.Col = SpaltennummerStatusBez Then
        If MSFlexGrid4.Text = "nicht bezahlt" Then
            MSFlexGrid4.Text = "bezahlt"
            
            sReNr = Trim(MSFlexGrid4.TextMatrix(MSFlexGrid4.Row, SpaltennummerReNr))
            speicher_statusbez sReNr, "bezahlt"
        Else
            MSFlexGrid4.Text = "nicht bezahlt"
            
            sReNr = Trim(MSFlexGrid4.TextMatrix(MSFlexGrid4.Row, SpaltennummerReNr))
            speicher_statusbez sReNr, "nicht bezahlt"
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid4_Click"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub speicher_statusbez(sReNr As String, sStatusbez As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    If sReNr = "" Then
        Exit Sub
    End If
    
    sSQL = "Update OFPO Set statusbez = '" & sStatusbez & "' "
    sSQL = sSQL & " where RENR = '" & sReNr & "'"
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_statusbez"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub speicher_bezInfo(sReNr As String, sbezInfo As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    If sReNr = "" Then
        Exit Sub
    End If
    
    sSQL = "Update OFPO Set zahlungsinfo = '" & sbezInfo & "' "
    sSQL = sSQL & " where RENR = '" & sReNr & "'"
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicher_bezInfo"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSFlexGrid4.Row
    lcol = MSFlexGrid4.Col
    
    If KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft Then
    
        Select Case lcol
            Case Is = SpaltennummerBezahlInfo
        
                If iKeypress = 0 And KeyCode <> vbKeyBack And KeyCode <> vbKeyF2 And KeyCode <> vbKeyReturn Then
                    MSFlexGrid4.Row = lrow
                    MSFlexGrid4.Col = lcol
                    MSFlexGrid4.Text = ""
                ElseIf iKeypress > 0 And KeyCode = 46 Then
                    MSFlexGrid4.Row = lrow
                    MSFlexGrid4.Col = lcol
                    MSFlexGrid4.Text = ""
                End If
                iKeypress = iKeypress + 1
        End Select
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid4_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSFlexGrid4_KeyPress(KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen    As String
    Dim cValid      As String
    Dim lcol        As Long
    Dim lrow        As Long
    
    lcol = MSFlexGrid4.Col
    lrow = MSFlexGrid4.Row
    
    lbl8.Caption = lrow

    cZeichen = Chr$(KeyAscii)
    
    Select Case lcol
        
        Case Is = SpaltennummerBezahlInfo
            gbAenderBezahlInfo = True
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß"
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If KeyAscii <> 0 Then
                MSFlexGrid4.Row = lrow
                MSFlexGrid4.Col = lcol
                cValid = MSFlexGrid4.Text
                If InStr(cValid, ",") > 0 And cZeichen = "," Then
                    KeyAscii = 0
                End If
                
                If KeyAscii <> 0 Then
                    If KeyAscii <> 8 Then
                        cValid = cValid & Chr$(KeyAscii)
                    Else
                        If Len(cValid) > 0 Then
                            cValid = Left$(cValid, Len(cValid) - 1)
                        End If
                    End If
                    MSFlexGrid4.Text = cValid
                End If
            End If
     End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid4_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSFlexGrid4_LeaveCell()
On Error GoTo LOKAL_ERROR
    
iKeypress = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid4_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid4_SelChange()
On Error GoTo LOKAL_ERROR

    Dim lColmerker  As Long
    Dim lRowmerker  As Long
    Dim sReNr       As String
    Dim sbezInfo    As String

    If gbAenderBezahlInfo Then
        lColmerker = MSFlexGrid4.Col
        lRowmerker = MSFlexGrid4.Row
        
        sbezInfo = MSFlexGrid4.TextMatrix(Val(lbl8.Caption), SpaltennummerBezahlInfo)
        sReNr = MSFlexGrid4.TextMatrix(Val(lbl8.Caption), SpaltennummerReNr)
        
        If Len(sbezInfo) > 35 Then
            MsgBox Trim(Left(sbezInfo, 35)) & " Achtung nur dieser Text wird gespeichert. (35 Zeichen)", vbInformation, "Winkiss Hinweis:"
        End If
        
        speicher_bezInfo sReNr, Trim(Left(sbezInfo, 35))

        MSFlexGrid4.Col = lColmerker
        MSFlexGrid4.Row = lRowmerker
        
        gbAenderBezahlInfo = False
        
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid4_SelChange"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus()
On Error GoTo LOKAL_ERROR

    Text1.BackColor = glSelBack1
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyReturn Then
        Command3_Click 12
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kreditverwaltung ist ein Fehler aufgetreten."
    
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
