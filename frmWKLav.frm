VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWKLav 
   BackColor       =   &H00C0C000&
   Caption         =   "Kundenanalyse"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKLav.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   1  'Fenstermitte
   Begin sevCommand3.Command Command2 
      Height          =   300
      Index           =   2
      Left            =   5280
      TabIndex        =   126
      Top             =   120
      Width           =   2895
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
      Caption         =   "Kundenbeteiligung"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   300
      Index           =   1
      Left            =   8400
      TabIndex        =   118
      Top             =   480
      Width           =   2895
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
      Caption         =   "Kundenduplikate ermitteln"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command2 
      Height          =   300
      Index           =   0
      Left            =   8400
      TabIndex        =   117
      Top             =   120
      Width           =   2895
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
      Caption         =   "Kundenanalyse (nicht gekauft)"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame fraSerienB 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   2535
      Left            =   5160
      TabIndex        =   100
      Top             =   5520
      Visible         =   0   'False
      Width           =   6855
      Begin sevCommand3.Command cmdSUeber 
         Height          =   375
         Left            =   240
         TabIndex        =   102
         Top             =   1920
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtSerienBHaupt 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   101
         Top             =   240
         Width           =   6255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00808000&
         Caption         =   "Text erstellen"
         Height          =   255
         Left            =   240
         TabIndex        =   103
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.Frame fraEmail 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      Height          =   2295
      Left            =   240
      TabIndex        =   96
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   1815
         Index           =   3
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   83
         Top             =   360
         Width           =   3735
      End
      Begin sevCommand3.Command cmdSenden 
         Height          =   375
         Left            =   120
         TabIndex        =   84
         Top             =   1800
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Senden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808000&
         Caption         =   "an Emailadresse"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   99
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808000&
         Caption         =   "Betreff"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   98
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808000&
         Caption         =   "Mitteilung"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   97
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame fraAusgabe 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   3240
      TabIndex        =   68
      Top             =   3120
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Frame fraSort 
         Appearance      =   0  '2D
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   3600
         TabIndex        =   110
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808000&
            Caption         =   "Umsatz absteigend"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   136
            Top             =   1560
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808000&
            Caption         =   "Postleitzahl"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   114
            Top             =   480
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808000&
            Caption         =   "Geburtsdatum"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   112
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00808000&
            Caption         =   "Nachname"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   111
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackColor       =   &H00808000&
            Caption         =   "sortiert nach"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   113
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.Frame fraFormat 
         Appearance      =   0  '2D
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   4680
         TabIndex        =   93
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   106
            Top             =   480
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "Steuerdatei erw."
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   105
            Top             =   0
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "Steuerdatei ein."
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   1440
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "als Word-Datei"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdFormat 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Top             =   960
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "als Excel-Datei"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin sevCommand3.Command cmdListen 
         Height          =   375
         Index           =   5
         Left            =   4800
         TabIndex        =   92
         Top             =   2040
         Width           =   1935
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "schließen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.Frame fraExport 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         Height          =   1455
         Left            =   2520
         TabIndex        =   85
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
         Begin sevCommand3.Command cmdExport 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   87
            Top             =   1080
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "in Datei"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdExport 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   600
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "per Email"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   80
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Frame fraEtiketten 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'Kein
         Height          =   2175
         Left            =   240
         TabIndex        =   77
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.TextBox txtFirmenbez 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   120
            MaxLength       =   100
            TabIndex        =   135
            Text            =   "Ihre Firmenbezeichnung"
            Top             =   50
            Width           =   1935
         End
         Begin sevCommand3.Command cmdEtikett 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   79
            ToolTipText     =   "Format: Zweckform 3475"
            Top             =   640
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "7,0 cm x 3,6 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdEtikett 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   78
            ToolTipText     =   "Format: Zweckform 3653"
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "10,5 cm x 4,24 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdEtikett 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   137
            Top             =   930
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "9,5 cm x 4,2 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdEtikett 
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   138
            ToolTipText     =   "Format: Zweckform 3424"
            Top             =   1210
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "10,5 cm x 4,8 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdEtikett 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   139
            ToolTipText     =   "Format: Zweckform L7160"
            Top             =   1500
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "6,35 cm x 3,81 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdEtikett 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   141
            ToolTipText     =   "Format: Zweckform 3652"
            Top             =   1780
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "7,0 cm x 4,23 cm"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin VB.Frame fraListen 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'Kein
         Height          =   2295
         Left            =   3960
         TabIndex        =   72
         Top             =   -360
         Visible         =   0   'False
         Width           =   2295
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   88
            Top             =   1920
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "Kundenliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "Bonusliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   75
            Top             =   960
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "Telefon/Fax Liste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   480
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "Geburtstagsliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command cmdListen 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   0
            Width           =   1935
            _ExtentX        =   0
            _ExtentY        =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            Caption         =   "Adressliste"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Serienbriefvorlage"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   720
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Adressetiketten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdAusgabe 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Listen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   2
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   6
         X1              =   6830
         X2              =   6830
         Y1              =   0
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   4
         X1              =   0
         X2              =   6840
         Y1              =   20
         Y2              =   20
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         BorderWidth     =   4
         Index           =   3
         X1              =   0
         X2              =   6840
         Y1              =   2510
         Y2              =   2510
      End
   End
   Begin Crystal.CrystalReport cr3 
      Left            =   11280
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   3
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileName   =   "C:\Thomas\VB Projekte\Winkiss\Datenbanken\niendorf\EXPORT\Mail.doc"
      PrintFileType   =   15
      WindowState     =   2
      EMailSubject    =   "Test"
      EMailMessage    =   "Hello"
      EMailToList     =   "thomasheinz@kisswws.de"
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   11280
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   2
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileName   =   "Kundenliste.doc"
      PrintFileType   =   17
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox lstdatnames 
      Height          =   840
      Left            =   2640
      TabIndex        =   63
      Top             =   7440
      Visible         =   0   'False
      Width           =   4815
   End
   Begin sevCommand3.Command cmdLaden 
      Height          =   375
      Left            =   480
      TabIndex        =   31
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Laden"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdStart 
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Ausführen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdDel 
      Height          =   375
      Left            =   7800
      TabIndex        =   28
      ToolTipText     =   "Löschen Ihrer Eingaben"
      Top             =   7920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Leeren"
      PictureAlign    =   2
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command cmdEnd 
      Height          =   375
      Left            =   9240
      TabIndex        =   27
      Top             =   7920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin sevCommand3.Command cmdPrint 
      Height          =   375
      Left            =   9240
      TabIndex        =   25
      Top             =   7440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Ausgabe"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   11280
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Winkiss Kundenanalyse"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ProgressBar pbrZeit 
      Height          =   375
      Left            =   6720
      TabIndex        =   29
      Top             =   6960
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame fraZuErstellen 
      BackColor       =   &H00C0C000&
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
      Height          =   6495
      Left            =   1320
      TabIndex        =   32
      Top             =   240
      Width           =   11055
      Begin VB.CheckBox chkDS 
         BackColor       =   &H00C0C000&
         Caption         =   "DS unterschrieben"
         Height          =   195
         Left            =   3480
         TabIndex        =   142
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtLief 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         MaxLength       =   6
         TabIndex        =   128
         Top             =   1920
         Width           =   975
      End
      Begin sevCommand3.Command Command4 
         Height          =   345
         Index           =   9
         Left            =   1320
         TabIndex        =   127
         Top             =   1920
         Width           =   360
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   285
         Index           =   1
         Left            =   10200
         TabIndex        =   115
         Top             =   4560
         Width           =   375
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtErtragBis 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox txtErtragVon 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   15
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox cboGebMonat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmWKLav.frx":0442
         Left            =   8640
         List            =   "frmWKLav.frx":046A
         TabIndex        =   23
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtUmsatzBis 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtUmsatzVon 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox checkOKr 
         BackColor       =   &H00C0C000&
         Caption         =   "offene Kredite"
         Height          =   195
         Left            =   9240
         TabIndex        =   14
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtFil 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7800
         MaxLength       =   2
         TabIndex        =   22
         Top             =   4080
         Width           =   615
      End
      Begin sevCommand3.Command cmdHinzu5 
         Height          =   255
         Left            =   8040
         TabIndex        =   66
         Top             =   4440
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "v"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox lstFil 
         Height          =   645
         Left            =   7800
         TabIndex        =   65
         Top             =   4800
         Width           =   615
      End
      Begin VB.ListBox lstLL 
         Height          =   840
         Left            =   240
         TabIndex        =   62
         Top             =   2640
         Width           =   4335
      End
      Begin VB.ListBox lstMerkmal 
         Height          =   645
         Left            =   6000
         TabIndex        =   61
         Top             =   4800
         Width           =   1215
      End
      Begin VB.ListBox lstOrt 
         Height          =   840
         Left            =   4920
         TabIndex        =   60
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ListBox lstAGN 
         Height          =   840
         Left            =   6960
         TabIndex        =   59
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox checkmannl 
         BackColor       =   &H00C0C000&
         Caption         =   "männlich"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox checkweibl 
         BackColor       =   &H00C0C000&
         Caption         =   "weiblich"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   670
         Width           =   975
      End
      Begin sevCommand3.Command cmdHinzu4 
         Height          =   255
         Left            =   7320
         TabIndex        =   58
         Top             =   2280
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "v"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdHinzu3 
         Height          =   255
         Left            =   2280
         TabIndex        =   57
         Top             =   2280
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "v"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtMerkmal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   21
         Top             =   4080
         Width           =   1215
      End
      Begin sevCommand3.Command cmdHinzu2 
         Height          =   255
         Left            =   6480
         TabIndex        =   56
         Top             =   4440
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "v"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command cmdHinzu1 
         Height          =   255
         Left            =   5640
         TabIndex        =   55
         Top             =   2280
         Width           =   255
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "v"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtKdNrVon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtKdNrBis 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtPlzVon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         MaxLength       =   7
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtOrt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         MaxLength       =   35
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtKauf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtKauf 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cboLin 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox cboAgn 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         TabIndex        =   13
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtBowertVon 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   17
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox txtBowertBis 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1031
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   18
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox txtDat1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtDat1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtDat2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox txtDat2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   4440
         Width           =   855
      End
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   20
         Left            =   6840
         TabIndex        =   129
         ToolTipText     =   "Kalender"
         Top             =   600
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
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   21
         Left            =   6840
         TabIndex        =   130
         ToolTipText     =   "Kalender"
         Top             =   960
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
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   30
         Left            =   8760
         TabIndex        =   131
         ToolTipText     =   "Kalender"
         Top             =   600
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   31
         Left            =   8760
         TabIndex        =   132
         ToolTipText     =   "Kalender"
         Top             =   960
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   0
         Left            =   5520
         TabIndex        =   133
         ToolTipText     =   "Kalender"
         Top             =   4080
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   360
         Index           =   1
         Left            =   5520
         TabIndex        =   134
         ToolTipText     =   "Kalender"
         Top             =   4440
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
         Image           =   20
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   345
         Index           =   2
         Left            =   10200
         TabIndex        =   140
         Top             =   2400
         Width           =   360
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
         Caption         =   "F2"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "alle Farben"
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
         Index           =   16
         Left            =   8640
         TabIndex        =   116
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ertrag"
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
         Index           =   15
         Left            =   840
         TabIndex        =   109
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         Height          =   255
         Left            =   240
         TabIndex        =   108
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   4080
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Geburtsmonat"
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
         Index           =   14
         Left            =   8880
         TabIndex        =   104
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   5
         X1              =   240
         X2              =   10560
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   10560
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         Height          =   255
         Left            =   9240
         TabIndex        =   91
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         Height          =   255
         Index           =   0
         Left            =   9240
         TabIndex        =   90
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Umsatz"
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
         Index           =   13
         Left            =   9720
         TabIndex        =   89
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Filiale"
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
         Index           =   12
         Left            =   7800
         TabIndex        =   67
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Geschlecht"
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
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kundennummer"
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
         Left            =   1920
         TabIndex        =   53
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Postleitzahl"
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
         Left            =   4080
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ortsnamen"
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
         Left            =   4920
         TabIndex        =   51
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kaufdatum"
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
         Left            =   7800
         TabIndex        =   50
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lieferant"
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
         Left            =   240
         TabIndex        =   49
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Linie"
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
         Index           =   6
         Left            =   3840
         TabIndex        =   48
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "AGN"
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
         Left            =   6960
         TabIndex        =   47
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         Height          =   255
         Left            =   1920
         TabIndex        =   46
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonuswert"
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
         Index           =   8
         Left            =   2880
         TabIndex        =   44
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         Height          =   255
         Left            =   7320
         TabIndex        =   43
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         Height          =   255
         Left            =   7320
         TabIndex        =   42
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   4155
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         Height          =   255
         Left            =   2400
         TabIndex        =   40
         Top             =   4500
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Merkmal"
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
         Index           =   9
         Left            =   6000
         TabIndex        =   39
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Geburtsdatum"
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
         Index           =   10
         Left            =   5880
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Datum 2"
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
         Index           =   11
         Left            =   4560
         TabIndex        =   37
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         Height          =   255
         Left            =   5400
         TabIndex        =   36
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "von:"
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   4155
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis:"
         Height          =   255
         Left            =   4080
         TabIndex        =   33
         Top             =   4500
         Width           =   375
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
      Height          =   5655
      Left            =   480
      TabIndex        =   64
      Top             =   960
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9975
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   12632256
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   6
      Left            =   8880
      TabIndex        =   123
      Top             =   3120
      Width           =   2295
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
      Caption         =   "zurücksetzen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   5
      Left            =   10800
      TabIndex        =   122
      ToolTipText     =   "Kalender"
      Top             =   2640
      Width           =   375
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
      Caption         =   "F2"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   3
      Left            =   8880
      TabIndex        =   121
      Top             =   4320
      Width           =   2295
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
      Caption         =   "alle zurücksetzen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   120
      Top             =   4800
      Width           =   2295
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
      Caption         =   "Verkaufsdetails"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   375
      Index           =   0
      Left            =   8880
      TabIndex        =   119
      Top             =   5280
      Width           =   2295
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
      Caption         =   "Kundendaten"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label18 
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
      Left            =   480
      TabIndex        =   125
      Top             =   6720
      Width           =   4575
   End
   Begin VB.Label lblAnzeige 
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
      Left            =   480
      TabIndex        =   30
      Top             =   7080
      Width           =   10815
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kunden - Analyse"
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
      Left            =   480
      TabIndex        =   26
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   11280
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "alle Farben"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   32
      Left            =   8880
      TabIndex        =   124
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frmWKLav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iPruef  As Integer
Dim bDat1   As Boolean
Dim bDat2   As Boolean
Dim bKauf   As Boolean
Dim bVorhanden As Boolean
Dim bEmail As Boolean
'Dim bDis As Boolean
Dim bDat As Boolean
Dim bExcel As Boolean
Dim bWord As Boolean

Dim lAusgewählt As Long

Dim sdateiname As String
Dim sErstelldatum As String
Dim bAender As Boolean
Dim bNotAll As Boolean
Dim bClickAusgabe As Boolean
Private Sub füllecboLinie(lLieferant As Long, cbox As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    sSQL = "SELECT LINBEZEICH,lpz FROM LINBEZ "
    sSQL = sSQL & " Where LINR = " & lLieferant & " "
    sSQL = sSQL & " order BY sorti "
    
    Set rs = gdBase.OpenRecordset(sSQL)
    
    cbox.Clear
   
    Do While Not rs.EOF
        cbox.AddItem rs!LPZ & Space(1) & rs!LINBEZEICH
        rs.MoveNext
    Loop
    
    rs.Close
    cbox.Text = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllecboLinie"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboAgn_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    cboAgn.BackColor = glSelBack1
    cboAgn.SelStart = 0
    cboAgn.SelLength = Len(cboAgn.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboAgn_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboAgn_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdHinzu4_Click
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboAgn_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboAgn_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    cboAgn.BackColor = vbWhite
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboAgn_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboGebMonat_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    cboGebMonat.BackColor = glSelBack1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboGebMonat_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboGebMonat_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    cboGebMonat.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboGebMonat_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

'Private Sub cboLief_Click()
'      On Error GoTo LOKAL_ERROR
'
'    Dim sLieferant As String
'
'        If cboLief.Text <> "" Then
'            sLieferant = cboLief.Text
'            füllecboLinie (sLieferant)
'        End If
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "cboLief_Click"
'    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub cboLief_GotFocus()
'    On Error GoTo LOKAL_ERROR
'
'    If bVorhanden Then
'        bAender = True
'    End If
'    cboLief.BackColor = glSelBack1
'    cboLief.SelStart = 0
'    cboLief.SelLength = Len(cboLief.Text)
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "cboLief_GotFocus"
'    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
'Private Sub cboLief_LostFocus()
'    On Error GoTo LOKAL_ERROR
'
'    cboLief.BackColor = vbWhite
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "cboLief_LostFocus"
'    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub
Private Sub cboLin_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    cboLin.BackColor = glSelBack1
    cboLin.SelStart = 0
    cboLin.SelLength = Len(cboLin.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboLin_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cboLin_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    cboLin.BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cboLin_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub checkmannl_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkmannl_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub checkOKr_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkOKr_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub checkweibl_GotFocus()
    On Error GoTo LOKAL_ERROR
    If bVorhanden Then
        bAender = True
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkweibl_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdAusgabe_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim rsrs As Recordset
    Dim sHaupt As String
    
    Select Case Index
        Case Is = 0
            fraListen.Visible = True
            fraEtiketten.Visible = False
            fraExport.Visible = False
            fraFormat.Visible = False
            fraSort.Visible = False
        Case Is = 1
            fraEtiketten.Visible = True
            fraSort.Visible = True
            fraListen.Visible = False
            fraExport.Visible = False
            fraFormat.Visible = False
            txtFirmenbez.SetFocus
        Case Is = 2
            'Serienbriefvorlage
            fraSerienB.Visible = True
            
            If Not NewTableSuchenDBKombi("haupt", gdBase) Then
            
            Else
                Set rsrs = gdBase.OpenRecordset("Haupt", dbOpenTable)
                If Not rsrs.RecordCount = 0 Then
                    sHaupt = rsrs!texthaupt
                    txtSerienBHaupt.Text = sHaupt
                End If
                rsrs.Close: Set rsrs = Nothing
            End If
            
            
            
                    
            fraEtiketten.Visible = False
            fraListen.Visible = False
            fraExport.Visible = False
            fraFormat.Visible = False
            fraSort.Visible = False
        Case Is = 3
            fraExport.Visible = True
            fraListen.Visible = False
            fraEtiketten.Visible = False
            fraFormat.Visible = False
            fraSort.Visible = False
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdAusgabe_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdEnd_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKLav
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEnd_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdDel_Click()
    On Error GoTo LOKAL_ERROR
    
    Unload frmWKLav
    frmWKLav.Show
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdDel_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtFirmenbez_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtFirmenbez.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtFirmenbez_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtFirmenbez_GotFocus()
On Error GoTo LOKAL_ERROR
    
    txtFirmenbez.BackColor = glSelBack1
    txtFirmenbez.SelStart = 0
    txtFirmenbez.SelLength = Len(txtFirmenbez.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtFirmenbez_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtLief_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtLief_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtLief_Change()
On Error GoTo LOKAL_ERROR

    If txtLief.Text <> "" Then
        füllecboLinie CLng(txtLief.Text), cboLin
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtLief_Change"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtLief_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtLief.BackColor = glSelBack1
    txtLief.SelStart = 0
    txtLief.SelLength = Len(txtLief.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtLief_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtLief_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtLief.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtLief_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtlief_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

If KeyCode = vbKeyF2 Then
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    
    
    gF2Prompt.bMultiple = False
    gF2Prompt.cFeld = "LINR"
    
    If gF2Prompt.cFeld <> "" Then
        frmWK00a.Show 1
    End If
    If gF2Prompt.cWahl <> "" Then
        txtLief.Text = gF2Prompt.cWahl
    End If
            
        
    txtLief.SetFocus
End If

    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtlief_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
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
        
        If Len(rs!AGTEXT) > 25 Then
            cboAgn.AddItem Left(rs!AGTEXT, 25) & "..." & rs!AGN
        Else
            cboAgn.AddItem rs!AGTEXT & Space(28 - Len(rs!AGTEXT)) & rs!AGN
            
        End If
        
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    cboAgn.Text = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllecboAgn"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdEtikett_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If NewTableSuchenDBKombi("KUTEILME", gdBase) Then
    
        loeschNEW "KUTTEN", gdBase
        CreateTableT2 "KUTTEN", gdBase
        
        sSQL = "Insert into KUTTEN Select  "
        sSQL = sSQL & " Knummer"
        sSQL = sSQL & ", KUERZEL"
        sSQL = sSQL & ", FIRMA"
        sSQL = sSQL & ", TITEL"
        sSQL = sSQL & ", NAME"
        sSQL = sSQL & ", VORNAME"
        sSQL = sSQL & ", STRASSE"
        sSQL = sSQL & ", PLZ"
        sSQL = sSQL & ", STADT"
        sSQL = sSQL & ", DATUM1"
        sSQL = sSQL & ", ANREDE"
        sSQL = sSQL & ", UMSATZ"
        
        sSQL = sSQL & " from KUTEILME "
        
        If Option1(0).Value = True Then
    '        Sortierung 1
            sSQL = sSQL & " order by Month(Datum1),Day(Datum1)"
        ElseIf Option1(1).Value = True Then
    '        Sortierung 2
            sSQL = sSQL & " order by Name"
        ElseIf Option1(2).Value = True Then
    '        Sortierung 3
            sSQL = sSQL & " order by Plz"
        ElseIf Option1(3).Value = True Then
    '        Sortierung 4
            sSQL = sSQL & " order by Umsatz desc"
        End If
        gdBase.Execute sSQL, dbFailOnError
        
        
        ' für den optionalen Firmentext
        
        If Trim(txtFirmenbez.Text) <> "Ihre Firmenbezeichnung" Then
            sSQL = "Update KUTTEN Set FIRMENBEZ = '" & Trim(txtFirmenbez.Text) & "'"
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        
        Select Case Index
            Case Is = 0
                'Etiketten groß
                If Modul6.FindFile(gcDBPfad, "aWKLavas.rpt") Then
                    reportbildschirm "spezial", "aWKLavas"
                Else
                    reportbildschirm "WKL017", "aWKLava"
                End If
                
               
            Case Is = 1
                'Etiketten klein
                
                If Modul6.FindFile(gcDBPfad, "aWKLavbs.rpt") Then
                    reportbildschirm "spezial", "aWKLavbs"
                Else
                    reportbildschirm "WKL017", "aWKLavb"
                End If
                
            Case Is = 2
                'Etiketten 9,5 x 4,2
'                If Modul6.FindFile(gcDBPfad, "aWKLavas.rpt") Then
'                    reportbildschirm "spezial", "aWKLavas"
'                Else
                    reportbildschirm "WKL017", "aWKLavi"
'                End If

            Case Is = 3
                'Etiketten 10,5 x 4,8
'                If Modul6.FindFile(gcDBPfad, "aWKLavas.rpt") Then
'                    reportbildschirm "spezial", "aWKLavas"
'                Else
                    reportbildschirm "WKL017", "aWKLavj"
'                End If

            Case Is = 4
                'Etiketten 6,35 x 3,81
'                If Modul6.FindFile(gcDBPfad, "aWKLavas.rpt") Then
'                    reportbildschirm "spezial", "aWKLavas"
'                Else
                    reportbildschirm "WKL017", "aWKLavk"
'                End If

            Case Is = 5
                'Etiketten klein
                
'                If Modul6.FindFile(gcDBPfad, "aWKLavbs.rpt") Then
'                    reportbildschirm "spezial", "aWKLavbs"
'                Else
                    reportbildschirm "WKL017", "aWKLavl"
'                End If
           
        End Select
    Else
        anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblanzeige
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdEtikett_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdExport_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            bEmail = True
'            bDis = False
            bDat = False
            fraFormat.Visible = True
'        Case Is = 1
'            bDis = True
'            bEmail = False
'            bDat = False
'            fraFormat.Visible = True
        Case Is = 2
            bDat = True
'            bDis = False
            bEmail = False
            fraFormat.Visible = True
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdExport_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdHinzu1_Click()
    On Error GoTo LOKAL_ERROR
    
    If txtOrt.Text <> "" Then
        lstOrt.AddItem (txtOrt.Text)
        txtOrt.Text = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdHinzu1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdHinzu2_Click()
    On Error GoTo LOKAL_ERROR
    
    If txtMerkmal.Text <> "" Then
        lstMerkmal.AddItem (txtMerkmal.Text)
        txtMerkmal.Text = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdHinzu2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdHinzu3_Click()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim lLinr As Long
    
    If txtLief.Text <> "" Then
        sSQL = " SELECT LINR, Liefbez FROM LISRT WHERE LINR = " & txtLief.Text & " "
        Set rs = gdBase.OpenRecordset(sSQL)
        
        If Not rs.EOF Then
            rs.MoveFirst
        
            If Not IsNull(rs!linr) Then
                lLinr = rs!linr
            Else
                Exit Sub
            End If
            
'            If Not IsNull(rs!Liefbez) Then
'                sLiefBez = rs!Liefbez
'            Else
'                sLiefBez = ""
'            End If
        End If
        rs.Close
        
        If cboLin.Text <> "" Then
            lstLL.AddItem (lLinr & "   " & Mid$(cboLin.Text, 1, InStr(1, cboLin.Text, " ")))
            cboLin.Text = ""
        Else
            lstLL.AddItem (lLinr & "      ")
        End If
        txtLief.Text = ""
        cboLin.Clear
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdHinzu3_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdHinzu4_Click()
    On Error GoTo LOKAL_ERROR
    
    If cboAgn.Text <> "" Then
        lstAGN.AddItem (Right(cboAgn.Text, 3))
        cboAgn.Text = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdHinzu4_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdHinzu5_Click()
    On Error GoTo LOKAL_ERROR
    
    If txtFil.Text <> "" Then
        lstFil.AddItem (txtFil.Text)
        txtFil.Text = ""
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdHinzu5_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdLaden_Click()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    Dim sdatname As String
    
    sSQL = "Select KAname , Bdate from KASQL order by Bdate desc, KAname"
    
    Set rs = gdBase.OpenRecordset(sSQL)
    
    lstdatnames.Clear
   
    Do While Not rs.EOF
        lstdatnames.AddItem rs!Bdate & "   " & rs!KAname
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    
    lstdatnames.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdLaden_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdListen_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Index <> 5 Then
        If NewTableSuchenDBKombi("KUTEILME", gdBase) = False Then
        
            anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblanzeige
            Exit Sub
        End If
    End If
        Select Case Index
            Case Is = 0     'Adressenliste
                reportbildschirm "kaali", "aWKLavc"
            Case Is = 1     'Geburtstagsliste
            
                sSQL = "Update KUTEILME set datum1 = '31.12.2010' where datum1 is null"
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Update KUTEILME set datum1 = '31.12.2010' where datum1 = ''"
                gdBase.Execute sSQL, dbFailOnError
            
                reportbildschirm "kagli", "aWKLavd"
            Case Is = 2     'Telefonliste
                reportbildschirm "katli", "aWKLave"
            Case Is = 3     'Bonusliste
                reportbildschirm "kaboli", "aWKLavf"
            Case Is = 4     'Kundenliste
                reportbildschirm "kakuli", "aWKLavg"
            Case Is = 5
                fraAusgabe.Visible = False
                fraEmail.Visible = False
                fraExport.Visible = False
                fraSort.Visible = False
                fraSerienB.Visible = False
                bClickAusgabe = False
        End Select
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdListen_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo LOKAL_ERROR
    
    If bClickAusgabe Then
        fraAusgabe.Visible = False
        bClickAusgabe = False
    Else
    
        KUTEILMEupdate
        
        If NewTableSuchenDBKombi("KUTEILME", gdBase) = False Then
            anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblanzeige
            Exit Sub
            
        Else
        
        

        End If
    
        fraAusgabe.Visible = True
        fraListen.Visible = False
        fraExport.Visible = False
        fraEtiketten.Visible = False
        fraEmail.Visible = False
        fraFormat.Visible = False
        fraSerienB.Visible = False
        
        bClickAusgabe = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdPrint_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdSenden_Click()
    On Error GoTo LOKAL_ERROR

    Dim sMailadress As String
    Dim sMessage As String
    Dim sBetreff As String
    Dim sPunkt As String
    Dim lPos As Long
    Dim lpos1 As Long
    Dim cPfad1 As String
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    lPos = InStr(Text1(0).Text, "@")
    
    If lPos > 0 Then
        sPunkt = Right(Text1(0).Text, Len(Text1(0)) - lPos)
        lpos1 = InStr(sPunkt, ".")
    End If
    
    
    
    If Text1(0).Text = "" Then
        Text1(0).SetFocus
        
        Exit Sub
        
    ElseIf lPos = 0 Then
        Text1(0).SetFocus
        
        Exit Sub
        
    ElseIf lpos1 <= 1 Then
        Text1(0).SetFocus
        
        Exit Sub
        
    ElseIf Right(Text1(0).Text, 1) = "." Then
        Text1(0).SetFocus
        
        Exit Sub
    
    Else
        If bExcel Then
        
            Dim Result      As String
            Dim Buff        As String
            Dim sZeitung    As String
            
            sZeitung = cPfad1 & "BOX\Kunden.xls"
        
            
            Buff = "mailto:" & Trim(Text1(0).Text)
            Buff = Buff & "?Subject=" & Trim(Text1(1).Text)
            Buff = Buff & "&Body=" & Trim(Text1(3).Text)
            Buff = Buff & "&Attach=" + Chr$(34) & sZeitung + Chr$(34)
            
        
            Result = ShellExecute(0&, "open", Buff, "", "", 6)
    
            
            
        ElseIf bWord Then
        
            cr2.ReportFileName = cPfad1 & "aWKLavg.rpt"
            cr2.PrintFileName = cPfad1 & "BOX\Kundenliste.doc"
            cr2.PrintFileType = crptRTF
            cr2.Destination = 3
            
            sMailadress = Text1(0).Text
            sBetreff = Text1(1).Text
            sMessage = Text1(3).Text
            
            cr2.EMailToList = sMailadress
            cr2.EMailMessage = sMessage
            cr2.EMailSubject = sBetreff
            cr2.Action = 1
            
            
        End If
        
        fraEmail.Visible = False
    End If
    
   
    
    
    
    bExcel = False
    bWord = False
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdSenden_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub cmdStart_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    Dim rs As Recordset
    
    lAusgewählt = 0

    If checkweibl.Value = 0 And checkmannl.Value = 0 And chkDS.Value = 0 And checkOKr.Value = 0 And txtKdNrVon.Text = "" And txtKdNrBis.Text = "" And _
    txtPlzVon.Text = "" And txtKauf(0).Text = "" And txtKauf(1).Text = "" And txtBowertVon.Text = "" And _
    txtBowertBis.Text = "" And txtUmsatzVon.Text = "" And txtUmsatzBis.Text = "" And txtDat1(0).Text = "" And _
    txtDat1(1).Text = "" And txtLief.Text = "" And Label1(16).Tag = "" And txtDat2(0).Text = "" And txtDat2(1).Text = "" And cboGebMonat.Text = "" And txtErtragVon.Text = "" And txtErtragBis.Text = "" Then
        bNotAll = True
    Else
        bNotAll = False
    End If
    
    If cboGebMonat.Text <> "" Then
        If Trim(cboGebMonat.Text) = "Januar" _
            Or Trim(cboGebMonat.Text) = "Februar" _
            Or Trim(cboGebMonat.Text) = "März" _
            Or Trim(cboGebMonat.Text) = "April" _
            Or Trim(cboGebMonat.Text) = "Mai" _
            Or Trim(cboGebMonat.Text) = "Juni" _
            Or Trim(cboGebMonat.Text) = "Juli" _
            Or Trim(cboGebMonat.Text) = "August" _
            Or Trim(cboGebMonat.Text) = "September" _
            Or Trim(cboGebMonat.Text) = "Oktober" _
            Or Trim(cboGebMonat.Text) = "November" _
            Or Trim(cboGebMonat.Text) = "Dezember" _
            Then
        Else
            MsgBox "Bitte wählen Sie einen Eintrag aus der Liste", vbOKOnly, "Winkiss Hinweis:"
            cboGebMonat.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    If bNotAll Then
        MsgBox "Bitte geben Sie mindestens ein Auswahlkriterium ein!", vbInformation, "Winkiss Eingabeaufforderung:"
    Else
        If bVorhanden Then
            If bAender Then
                iRet = (MsgBox("Wollen Sie die Veränderung speichern?", vbQuestion + vbYesNo, "Winkiss Frage:"))
                If iRet = vbYes Then
                    Zusammenstellungerstellen Trim(sdateiname)
                Else
                
                End If
            End If
        Else
            sdateiname = InputBox("Wollen Sie diese Zusammenstellung speichern?. Dann vergeben Sie bitte einen Namen!", "Winkiss Frage:")
            sSQL = "select * from KASQL where KANAME = '" & sdateiname & "' "
            Set rs = gdBase.OpenRecordset(sSQL)
        
            If Not rs.EOF Then
                Do
                    rs.Close: Set rs = Nothing
                    sdateiname = InputBox("Der Name ist schon vergeben. Bitte wählen Sie einen neuen Namen aus!", "Winkiss Eingabe:")
                    sSQL = "select * from KASQL where KANAME = '" & sdateiname & "' "
                    Set rs = gdBase.OpenRecordset(sSQL)
                Loop Until rs.EOF
            End If
            rs.Close: Set rs = Nothing
        
            If sdateiname = "" Then
                sdateiname = "kein Betreff"
            End If
        Zusammenstellungerstellen sdateiname
    End If
        
        Screen.MousePointer = 11
        
        ausführen Trim(sdateiname), sErstelldatum
        Zusammenstellunganzeigen
        
        Screen.MousePointer = 0

        bVorhanden = False
        bAender = False
        bNotAll = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdStart_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zusammenstellunganzeigen()
    On Error GoTo LOKAL_ERROR
    
    Tabelleerstellen
    
    If NewTableSuchenDBKombi("KUTEILME", gdBase) Then
        Tabellefuellen
        
        Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zusammenstellunganzeigen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSHFlexGrid, siEigFak As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim siFak       As Single
    Dim bBreit()    As Long
    Dim i           As Long
    Dim j           As Long
    
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
Private Sub Tabelleerstellen()
    On Error GoTo LOKAL_ERROR

    
    
    With MSHFLEX1
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 8
        .FixedCols = 0
        .FixedRows = 1
        
        .Row = 0
        
        .Col = 0
        .ColWidth(0) = 620
        .Text = "OK"
   
        
        .Col = 1
        .ColWidth(1) = 800
        .Text = "Kundennr"
        
        .Col = 2
        .ColWidth(2) = 1500
        .Text = "Vorname"
        
        .Col = 3
        .ColWidth(3) = 1600
        .Text = "Name"
        
        .Col = 4
        .ColWidth(4) = 1600
        .Text = "Straße"
        
        .Col = 5
        .ColWidth(5) = 600
        .Text = "Plz"
        
        .Col = 6
        .ColWidth(6) = 1600
        .Text = "Ort"
        
        .Col = 7
        .ColWidth(7) = 1000
        .Text = "Geburtstag"

    
    End With
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabelleerstellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellefuellen()
    On Error GoTo LOKAL_ERROR

    Dim rsKUTEILME      As Recordset
    Dim lrow            As Long
    Dim lWert           As Long
    Dim sWert           As String
    Dim lCounter        As Long
    
    
    Set rsKUTEILME = gdBase.OpenRecordset("KUTEILME", dbOpenTable)
    
    lrow = 1
    If Not rsKUTEILME.EOF Then
        rsKUTEILME.MoveFirst
        
        MSHFLEX1.Redraw = False
        
        anzeige "normal", "Kunden werden ermittelt...", lblanzeige
        
        pbrZeit.Visible = True
        pbrZeit.Max = 300
        
        Do While Not rsKUTEILME.EOF
            
            lrow = lrow + 1
            lCounter = lCounter + 1
            
            If lCounter = 300 Then
                lCounter = 0
            End If
            pbrZeit.Value = lCounter
            
            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Row = lrow
            
            MSHFLEX1.Col = 0
            MSHFLEX1.Text = "X"
            
            If Not IsNull(rsKUTEILME!knummer) Then
                lWert = rsKUTEILME!knummer
            Else
                lWert = 0
            End If
            
            MSHFLEX1.Col = 1
            MSHFLEX1.Text = lWert
            
            Dim sKUNDNR     As String
            Dim cAWM        As String
            
            
            sKUNDNR = lWert
            cAWM = ""
            If sKUNDNR <> "" Then
                cAWM = WhatIsAwmKU(sKUNDNR)
            Else
                
            End If
            
            If cAWM = "" Then cAWM = "0"
            FaerbenFlexHKunde cAWM, MSHFLEX1, 1, lrow
            
            If Not IsNull(rsKUTEILME!vorname) Then
                sWert = rsKUTEILME!vorname
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 2
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!name) Then
                sWert = rsKUTEILME!name
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 3
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!strasse) Then
                sWert = rsKUTEILME!strasse
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 4
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!Plz) Then
                sWert = rsKUTEILME!Plz
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 5
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!STADT) Then
                sWert = rsKUTEILME!STADT
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 6
            MSHFLEX1.Text = Trim(sWert)
            
            If Not IsNull(rsKUTEILME!Datum1) Then
                sWert = rsKUTEILME!Datum1
            Else
                sWert = ""
            End If
            
            MSHFLEX1.Col = 7
            MSHFLEX1.Text = Trim$(sWert)
    
            rsKUTEILME.MoveNext
        Loop
        pbrZeit.Visible = False
    End If
    rsKUTEILME.Close
    
    MSHFLEX1.RowHeight(1) = 0
    lrow = lrow - 1
    
    lAusgewählt = lrow
    
    If lrow > 1 Then
        anzeige "normal", lrow & " Kunden wurden ermittelt.", lblanzeige
        anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label18
    ElseIf lrow = 1 Then
        anzeige "normal", lrow & " Kunde wurden ermittelt.", lblanzeige
        anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label18
    Else
        anzeige "rot", "Es wurden keine Kunden ermittelt.", lblanzeige
        anzeige "normal", "", Label18
        
        pbrZeit.Visible = False
        Exit Sub
    End If
    
    fraZuErstellen.Visible = False
    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellefuellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenvorhanden()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim sPfad As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If


    If Not tableSuchenDBKombi("KASQL", 1) Then
        CreateTable "KASQL", gdBase
        
    End If
    

    If Not tableSuchenDBKombi("KASQLAGN", 1) Then
    
        sSQL = "Create Table KASQLAGN "
        sSQL = sSQL & " ( "
        sSQL = sSQL & " KAName Text(50) "
        sSQL = sSQL & ", AGN integer) "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    End If
    
'    If Not Modul6.FindFile(sPfad, "KASQLORT.DBF") Then
    If Not tableSuchenDBKombi("KASQLORT", 1) Then
        sSQL = "Create Table KASQLORT "
        sSQL = sSQL & " ( "
        sSQL = sSQL & " KAName Text(50) "
        sSQL = sSQL & ", ORT Text(35)) "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    End If
    
'    If Not Modul6.FindFile(sPfad, "KASQLMK.DBF") Then
    If Not tableSuchenDBKombi("KASQLMK", 1) Then
        sSQL = "Create Table KASQLMK "
        sSQL = sSQL & " ( "
        sSQL = sSQL & " KAName Text(50) "
        sSQL = sSQL & ", Merkmal Text(10)) "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    End If
    
'    If Not Modul6.FindFile(sPfad, "KASQLLL.DBF") Then
    If Not tableSuchenDBKombi("KASQLLL", 1) Then
        sSQL = "Create Table KASQLLL "
        sSQL = sSQL & " ( "
        sSQL = sSQL & " KAName Text(50) "
        sSQL = sSQL & ", Lieferant long "
        sSQL = sSQL & ", Linie Text(3) ) "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    End If
    
'    If Not Modul6.FindFile(sPfad, "KASQLFIL.DBF") Then
    If Not tableSuchenDBKombi("KASQLFIL", 1) Then
        sSQL = "Create Table KASQLFIL "
        sSQL = sSQL & " ( "
        sSQL = sSQL & " KAName Text(50) "
        sSQL = sSQL & ", Filiale byte) "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Tabellenvorhanden"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zusammenstellungerstellen(sdatname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim sPfad As String
    
    Dim sGeschlecht As String
    Dim lKdNumVon As Long
    Dim lKdnumBis As Long
    Dim sPlzVon As String
    Dim sawm As String
    
    Dim sKaufdatVon As String
    Dim sKaufdatBis As String
    Dim dBowertVon As Double
    Dim dBowertBis As Double
    Dim dErtragVon As Double
    Dim dErtragBis As Double
    Dim dUmsatzVon As Double
    Dim dUmsatzBis As Double
    Dim sDat1Von As String
    Dim sDat1Bis As String
    Dim sDat2Von As String
    Dim sDat2Bis As String
    
    
    Dim iAGN As Integer
    Dim sOrt As String
    Dim sMerkmal As String
    Dim lLieferant As Long
    Dim sLinie As String
    Dim byFil As Byte
    Dim sKreditfrage As String
    Dim iGebMonat As Integer
    Dim bDS As Boolean
    
    Dim i As Integer
    
    
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    
    sGeschlecht = ""
    
    If checkmannl.Value = 1 And checkweibl.Value = 1 Then
        checkmannl.Value = 0
        checkweibl.Value = 0
    End If
    
    
    
    If checkweibl.Value = 1 Then
        sGeschlecht = "W"
    ElseIf checkmannl.Value = 1 Then
        sGeschlecht = "M"
    Else
        sGeschlecht = ""
    End If
    
    bDS = False
    If chkDS.Value = vbChecked Then
        bDS = True
    End If
    
    
    
    
    If checkOKr.Value = 1 Then
        sKreditfrage = "J"
    Else
        sKreditfrage = ""
    End If
    
    If txtKdNrVon.Text <> "" Then
        If txtKdNrVon.Text = "0" Then txtKdNrVon.Text = "1"
        lKdNumVon = txtKdNrVon.Text
    End If
    
    If txtKdNrBis.Text <> "" Then
        lKdnumBis = txtKdNrBis.Text
    End If

    If txtPlzVon.Text <> "" Then
        sPlzVon = txtPlzVon.Text
    End If
    
    If Label1(16).Tag <> "" Then
        sawm = Label1(16).Tag
    End If
    
    If txtKauf(0).Text <> "" Then
        sKaufdatVon = txtKauf(0).Text
        sKaufdatBis = DateValue(Now)
    End If
    
    If txtKauf(1).Text <> "" Then
        sKaufdatBis = txtKauf(1).Text
    End If
    
    If txtBowertVon.Text <> "" Then
        dBowertVon = txtBowertVon.Text
    End If
    
    If txtBowertBis.Text <> "" Then
        dBowertBis = txtBowertBis.Text
    End If
    
    If txtErtragVon.Text <> "" Then
        dErtragVon = txtErtragVon.Text
    End If
    
    If txtErtragBis.Text <> "" Then
        dErtragBis = txtErtragBis.Text
    End If
    
    If txtUmsatzVon.Text <> "" Then
        dUmsatzVon = txtUmsatzVon.Text
    End If
    
    If txtUmsatzBis.Text <> "" Then
        dUmsatzBis = txtUmsatzBis.Text
    End If
    
    If txtDat1(0).Text <> "" Then
        sDat1Von = txtDat1(0).Text
        sDat1Bis = DateValue(Now)
    End If

    If txtDat1(1).Text <> "" Then
        sDat1Bis = txtDat1(1).Text
    End If
    
    If txtDat2(0).Text <> "" Then
        sDat2Von = txtDat2(0).Text
        sDat2Bis = DateValue(Now)
    End If
    
    If txtDat2(1).Text <> "" Then
        sDat2Bis = txtDat2(1).Text
    End If
    
    If cboGebMonat.Text <> "" Then
        Select Case cboGebMonat.Text
            Case Is = "Januar"
                iGebMonat = 1
            Case Is = "Februar"
                iGebMonat = 2
            Case Is = "März"
                iGebMonat = 3
            Case Is = "April"
                iGebMonat = 4
            Case Is = "Mai"
                iGebMonat = 5
            Case Is = "Juni"
                iGebMonat = 6
            Case Is = "Juli"
                iGebMonat = 7
            Case Is = "August"
                iGebMonat = 8
            Case Is = "September"
                iGebMonat = 9
            Case Is = "Oktober"
                iGebMonat = 10
            Case Is = "November"
                iGebMonat = 11
            Case Is = "Dezember"
                iGebMonat = 12
            Case Else
                iGebMonat = 0
        End Select

    End If
    
    
    sSQL = "Delete from KASQL where KANAME = '" & sdatname & "' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLLL where KAname = '" & sdatname & "' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLMK where KAname = '" & sdatname & "' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLORT where KAname = '" & sdatname & "' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLAGN where KAname = '" & sdatname & "' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLFIL where KAname = '" & sdatname & "' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KASQL "
    sSQL = sSQL & " (KAName, Bdate, Geschlecht, KdNumVon, KdnumBis, PlzVon,awm, KaufdatVon, KaufdatBis "
    sSQL = sSQL & ", BowertVon, BowertBis , ErtragVon, ErtragBis, UmsatzVon , UmsatzBis, Dat1Von, Dat1Bis, Dat2Von, Dat2Bis ,Kredit, Gebmonat,DS"
    sSQL = sSQL & " )"
    
    sSQL = sSQL & " VALUES ( "
    sSQL = sSQL & "'" & sdatname & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & DateValue(Now) & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sGeschlecht & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " " & lKdNumVon & " "
    sSQL = sSQL & " , "
    sSQL = sSQL & " " & lKdnumBis & " "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sPlzVon & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sawm & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sKaufdatVon & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sKaufdatBis & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " '" & dBowertVon & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " '" & dBowertBis & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " '" & dErtragVon & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " '" & dErtragBis & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " '" & dUmsatzVon & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " '" & dUmsatzBis & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sDat1Von & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sDat1Bis & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sDat2Von & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sDat2Bis & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & "'" & sKreditfrage & "' "
    sSQL = sSQL & " , "
    sSQL = sSQL & " " & iGebMonat & " "
    
    sSQL = sSQL & " , "
    
    If bDS = False Then
        sSQL = sSQL & "  False "
    Else
        sSQL = sSQL & " TRUE "
    End If
    
    
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    cmdHinzu4_Click
    If Not lstAGN.ListCount < 0 Then
        For i = 0 To lstAGN.ListCount - 1
            
            iAGN = lstAGN.list(i)
            
            sSQL = "Insert into KASQLAGN "
            sSQL = sSQL & " (KAName, AGN )"
            sSQL = sSQL & " VALUES ( "
            sSQL = sSQL & "'" & sdatname & "' "
            sSQL = sSQL & " , "
            sSQL = sSQL & "" & iAGN & " "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            
        Next i
    End If
    
    cmdHinzu5_Click
    If Not lstFil.ListCount < 0 Then
        For i = 0 To lstFil.ListCount - 1
            
            byFil = lstFil.list(i)
            
            sSQL = "Insert into KASQLFIL "
            sSQL = sSQL & " (KAName, Filiale )"
            sSQL = sSQL & " VALUES ( "
            sSQL = sSQL & "'" & sdatname & "' "
            sSQL = sSQL & " , "
            sSQL = sSQL & "" & byFil & " "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            
        Next i
    End If
    
    cmdHinzu1_Click
    If Not lstOrt.ListCount < 0 Then
        For i = 0 To lstOrt.ListCount - 1
            
            sOrt = lstOrt.list(i)
            
            sSQL = "Insert into KASQLOrt "
            sSQL = sSQL & " (KAName, Ort )"
            sSQL = sSQL & " VALUES ( "
            sSQL = sSQL & "'" & sdatname & "' "
            sSQL = sSQL & " , "
            sSQL = sSQL & "'" & sOrt & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            
        Next i
    End If
    
    cmdHinzu2_Click
    If Not lstMerkmal.ListCount < 0 Then
        For i = 0 To lstMerkmal.ListCount - 1
            
            sMerkmal = lstMerkmal.list(i)
            
            sSQL = "Insert into KASQLMK "
            sSQL = sSQL & " (KAName, Merkmal )"
            sSQL = sSQL & " VALUES ( "
            sSQL = sSQL & "'" & sdatname & "' "
            sSQL = sSQL & " , "
            sSQL = sSQL & "'" & sMerkmal & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
            
        Next i
    End If
    
    cmdHinzu3_Click
    If Not lstLL.ListCount < 0 Then
        For i = 0 To lstLL.ListCount - 1
            
            lLieferant = Left(lstLL.list(i), 6)
            sLinie = Right(lstLL.list(i), 3)
            
            sSQL = "Insert into KASQLLL "
            sSQL = sSQL & " (KAName, Lieferant, Linie )"
            sSQL = sSQL & " VALUES ( "
            sSQL = sSQL & "'" & sdatname & "' "
            sSQL = sSQL & " , "
            sSQL = sSQL & "" & lLieferant & " "
            sSQL = sSQL & " , "
            sSQL = sSQL & "'" & sLinie & "' "
            sSQL = sSQL & " ) "
            gdBase.Execute sSQL, dbFailOnError
        Next i
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zusammenstellungerstellen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ausführen(sdatname As String, sErstelldat As String)
    On Error GoTo LOKAL_ERROR

    Dim sPfad           As String
    Dim sSQL            As String
    Dim sSQLKunden      As String
    Dim sSQLAGN         As String
    
    Dim rsKd            As Recordset
    
    Dim rsKASQL         As Recordset
    Dim rsKASQLLL       As Recordset
    Dim rsKASQLMK       As Recordset
    Dim rsKASQLAGN      As Recordset
    Dim rsKASQLORT      As Recordset
    Dim rsKASQLFIL      As Recordset
    Dim sSQLLL          As String
    
    Dim sGeschlecht     As String
    Dim lKdNumVon       As Long
    Dim lKdnumBis       As Long
    Dim sawm            As String
    Dim sPlzVon         As String
    Dim sKaufdatVon     As String
    Dim sKaufdatBis     As String
    Dim dBowertVon      As Double
    Dim dBowertBis      As Double
    Dim dErtragVon      As Double
    Dim dErtragBis      As Double
    Dim dUmsatzVon      As Double
    Dim dUmsatzBis      As Double
    Dim sDat1Von        As String
    Dim sDat1Bis        As String
    Dim sDat2Von        As String
    Dim sDat2Bis        As String
    Dim sKredit         As String
  
    
    Dim iAGN            As Integer
    Dim sOrt            As String
    Dim sMerkmal        As String
    Dim lLieferant      As Long
    Dim sLinie          As String
    Dim byFil           As Byte
    
    Dim lKaufdatVon     As Long
    Dim lKaufdatBis     As Long
    Dim lDat            As Long
    Dim iGebMonat       As Integer
    Dim bDS             As Boolean
    
    Dim iCount, l, k As Integer
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    
    anzeige "Normal", "Kunden werden ermittelt...", lblanzeige
    
    pbrZeit.Visible = True
    pbrZeit.Max = 300
    
    sSQL = "Select * from KASQL "
    sSQL = sSQL & " where KAname = "
    sSQL = sSQL & "'" & sdatname & "' "
    
    Set rsKASQL = gdBase.OpenRecordset(sSQL)
        
    If Not rsKASQL.EOF Then
        rsKASQL.MoveFirst
        
        If Not IsNull(rsKASQL!DS) Then
            If rsKASQL!DS = True Then
                bDS = True
            Else
                bDS = False
            End If
        End If
    
        If Not IsNull(rsKASQL!geschlecht) Then
            sGeschlecht = rsKASQL!geschlecht
        Else
            sGeschlecht = ""
        End If
        
        If Not IsNull(rsKASQL!KREDIT) Then
            sKredit = rsKASQL!KREDIT
        Else
            sKredit = ""
        End If
        
        If Not IsNull(rsKASQL!Gebmonat) Then
            iGebMonat = rsKASQL!Gebmonat
        Else
            iGebMonat = "0"
        End If
        
        If Not IsNull(rsKASQL!KdNumVon) Then
            lKdNumVon = rsKASQL!KdNumVon
        Else
            lKdNumVon = "0"
        End If
        
        If Not IsNull(rsKASQL!KdnumBis) Then
            lKdnumBis = rsKASQL!KdnumBis
        Else
            lKdnumBis = 0
        End If
        
        If lKdnumBis = 0 And lKdNumVon <> 0 Then
            lKdnumBis = rsKASQL!KdNumVon
        End If
        
        If Not IsNull(rsKASQL!PlzVon) Then
            sPlzVon = rsKASQL!PlzVon
        Else
            sPlzVon = ""
        End If
        
        If Not IsNull(rsKASQL!AWM) Then
            sawm = rsKASQL!AWM
        Else
            sawm = ""
        End If
        
        If rsKASQL!KaufdatVon <> "" Then
            sKaufdatVon = rsKASQL!KaufdatVon
            lKaufdatVon = DateValue(sKaufdatVon)
            sKaufdatVon = CLng(lKaufdatVon)
        Else
            sKaufdatVon = ""
        End If
        
        If rsKASQL!KaufdatBis <> "" Then
            sKaufdatBis = rsKASQL!KaufdatBis
            lKaufdatBis = DateValue(sKaufdatBis)
            sKaufdatBis = CLng(lKaufdatBis)
        Else
            sKaufdatBis = ""
        End If
            
        If sKaufdatBis = "" Then
            lKaufdatBis = DateValue(Now)
            sKaufdatBis = CLng(lKaufdatBis)
        End If
        '****** Bowert
        If Not IsNull(rsKASQL!BowertVon) Then
            dBowertVon = Format$(rsKASQL!BowertVon, "######0")
        Else
            dBowertVon = 0
        End If
        
        If Not IsNull(rsKASQL!BowertBis) Then
            dBowertBis = Format$(rsKASQL!BowertBis, "######0")
        Else
            dBowertBis = 0
        End If
        
        If dBowertBis = 0 Then
            sSQL = "select Max(Bonus)as MoM from Kunden"
            Set rsKd = gdBase.OpenRecordset(sSQL)
            
            If Not rsKd.EOF Then
                rsKd.MoveFirst
        
                If Not IsNull(rsKd!MoM) Then
                    dBowertBis = Format$(rsKd!MoM, "######0")
                Else
                    dBowertBis = 0
                End If
            End If
            rsKd.Close
        End If
        
        '****** Ertrag
        If Not IsNull(rsKASQL!ErtragVon) Then
            dErtragVon = Format$(rsKASQL!ErtragVon, "######0")
        Else
            dErtragVon = 0
        End If
        
        If Not IsNull(rsKASQL!ErtragBis) Then
            dErtragBis = Format$(rsKASQL!ErtragBis, "######0")
        Else
            dErtragBis = 0
        End If
        
        '******Ende Ertrag
            
        If Not IsNull(rsKASQL!UmsatzVon) Then
            dUmsatzVon = Format$(rsKASQL!UmsatzVon, "######0")
        Else
            dUmsatzVon = 0
        End If
        
        If Not IsNull(rsKASQL!UmsatzBis) Then
            dUmsatzBis = Format$(rsKASQL!UmsatzBis, "######0")
        Else
            dUmsatzBis = 0
        End If
        
        If dUmsatzBis = 0 And dUmsatzVon <> 0 Then
            dUmsatzBis = Format$(rsKASQL!UmsatzVon, "######0")
        End If
        
        If rsKASQL!dat1Von <> "" Then
            sDat1Von = Format(rsKASQL!dat1Von, "DD.MM.YYYY")
            
            lDat = DateValue(sDat1Von)
            sDat1Von = CLng(lDat)
        Else
            sDat1Von = ""
        End If
        
        If rsKASQL!dat1Bis <> "" Then
            sDat1Bis = Format(rsKASQL!dat1Bis, "DD.MM.YYYY")
            
            lDat = DateValue(sDat1Bis)
            sDat1Bis = CLng(lDat)
        Else
            sDat1Bis = ""
        End If
        
        If sDat1Bis = "" Then
            lDat = DateValue(Now)
            sDat1Bis = CLng(lDat)
        End If
        
        If rsKASQL!dat2Von <> "" Then
            sDat2Von = rsKASQL!dat2Von
            lDat = DateValue(sDat2Von)
            sDat2Von = CLng(lDat)
        Else
            sDat2Von = ""
        End If
        
        If rsKASQL!dat2Bis <> "" Then
            sDat2Bis = rsKASQL!dat2Bis
            lDat = DateValue(sDat2Bis)
            sDat2Bis = CLng(lDat)
        Else
            sDat2Bis = ""
        End If
        
        If sDat2Bis = "" Then
            lDat = DateValue(Now)
            sDat2Bis = CLng(lDat)
        End If
        
    End If
    rsKASQL.Close
    
    pbrZeit.Value = 100
    loeschNEW "KUTEILME", gdBase
    
    sSQLKunden = " Select   "
    sSQLKunden = sSQLKunden & " Kunden.KUNDNR as Knummer"
    sSQLKunden = sSQLKunden & ", Kunden.KUERZEL "
    sSQLKunden = sSQLKunden & ", Kunden.FIRMA "
    sSQLKunden = sSQLKunden & ", Kunden.TITEL "
    sSQLKunden = sSQLKunden & ", Kunden.NAME "
    sSQLKunden = sSQLKunden & ", Kunden.VORNAME "
    sSQLKunden = sSQLKunden & ", Kunden.STRASSE "
    sSQLKunden = sSQLKunden & ", Kunden.PLZ "
    sSQLKunden = sSQLKunden & ", Kunden.STADT "
    sSQLKunden = sSQLKunden & ", Kunden.TEL "
    sSQLKunden = sSQLKunden & ", Kunden.FAXNR "
    sSQLKunden = sSQLKunden & ", Kunden.MERKMAL "
    sSQLKunden = sSQLKunden & ", Kunden.ANREDE "
    sSQLKunden = sSQLKunden & ", Kunden.MERKMAL2 "
    sSQLKunden = sSQLKunden & ", Kunden.FORMATDAT "
    sSQLKunden = sSQLKunden & ", Kunden.RECHNR "
    sSQLKunden = sSQLKunden & ", Kunden.KURZTEXT1 "
    sSQLKunden = sSQLKunden & ", Kunden.KURZTEXT2 "
    sSQLKunden = sSQLKunden & ", format(Kunden.DATUM1,'DD.MM.YY') as datum1 "
    sSQLKunden = sSQLKunden & ", format(Kunden.DATUM2,'DD.MM.YY') as datum2 "
    sSQLKunden = sSQLKunden & ", Kunden.UMSLJ "
    sSQLKunden = sSQLKunden & ", Kunden.UMSVJ "
    sSQLKunden = sSQLKunden & ", Kunden.OSUM "
    sSQLKunden = sSQLKunden & ", Kunden.KASSE "
    sSQLKunden = sSQLKunden & ", Kunden.RABATT "
    sSQLKunden = sSQLKunden & ", Kunden.FILIALNR "
    sSQLKunden = sSQLKunden & ", Kunden.GESCHLECHT "
    sSQLKunden = sSQLKunden & ", Kunden.ECIDENT "
    sSQLKunden = sSQLKunden & ", Kunden.GESPERRT "
    sSQLKunden = sSQLKunden & ", Kunden.KUNDKART "
    sSQLKunden = sSQLKunden & ", Kunden.BONUS "
    sSQLKunden = sSQLKunden & ", Kunden.PREISKZ "
    sSQLKunden = sSQLKunden & ", Kunden.Angelegt "
    sSQLKunden = sSQLKunden & ", Kunden.Aender "
    sSQLKunden = sSQLKunden & ", Kunden.Lastdate "
    sSQLKunden = sSQLKunden & ", Kunden.Lasttime "
    sSQLKunden = sSQLKunden & ", Kunden.EMAIL "
    sSQLKunden = sSQLKunden & ", Kunden.MOBILTEL "
    sSQLKunden = sSQLKunden & ", Kunden.awm "
    sSQLKunden = sSQLKunden & ", Kunden.DS "
    sSQLKunden = sSQLKunden & ", ' ' as OKredit "
    sSQLKunden = sSQLKunden & ", '" & sdatname & "' as Datname "
    sSQLKunden = sSQLKunden & ", '" & sErstelldat & "' as Daterstellung "
    sSQLKunden = sSQLKunden & ", 0.00 as Ertrag"
    sSQLKunden = sSQLKunden & ", 0.00 as Umsatz"

    sSQLKunden = sSQLKunden & " into KUTEILME from Kunden "
    pbrZeit.Value = 150
    '************************Lieferant/Linie
        
        sSQL = "Select * From KASQLLL Where KANAME = '" & sdatname & "'"
        Set rsKASQLLL = gdBase.OpenRecordset(sSQL)

        If Not rsKASQLLL.EOF Then
        
            sSQLLL = sSQLLL & " and ("
            rsKASQLLL.MoveLast
            
            iCount = rsKASQLLL.RecordCount
            rsKASQLLL.MoveFirst
            
            For l = 0 To iCount - 1

            
                If Not IsNull(rsKASQLLL!Lieferant) Then
                    lLieferant = rsKASQLLL!Lieferant
                Else
                    lLieferant = "0"
                End If
                
                If Trim(rsKASQLLL!Linie) <> "" Then
                    sLinie = rsKASQLLL!Linie
                Else
                    sLinie = ""
                End If
                
                If sLinie <> "" Then
                    sSQLLL = sSQLLL & " ( "
                End If
                
                sSQLLL = sSQLLL & " kassjour.LINR = " & lLieferant & " "
                
                If sLinie <> "" Then
                    sSQLLL = sSQLLL & " and "
                    sSQLLL = sSQLLL & " kassjour.LPZ = " & sLinie & " "
                    sSQLLL = sSQLLL & " ) "
                End If
                
                If l = iCount - 1 Then
                    sSQLLL = sSQLLL & " ) "
                Else
                    sSQLLL = sSQLLL & " or "
                End If
                
                rsKASQLLL.MoveNext
            Next l
        End If
        rsKASQLLL.Close
    
    '************************AGN
        
        sSQL = "Select * From KASQLAGN Where KANAME = '" & sdatname & "'"
        Set rsKASQLAGN = gdBase.OpenRecordset(sSQL)

        If Not rsKASQLAGN.EOF Then
        
            sSQLAGN = sSQLAGN & " and ("
            rsKASQLAGN.MoveLast
            
            iCount = rsKASQLAGN.RecordCount
            rsKASQLAGN.MoveFirst
            
            For l = 0 To iCount - 1

'                If Trim(rsKASQLAGN!AGN) <> "" Then
                If Not IsNull(rsKASQLAGN!AGN) Then
                    iAGN = rsKASQLAGN!AGN
                Else
                    iAGN = "0"
                End If
        
                sSQLAGN = sSQLAGN & " kassjour.AGN = " & iAGN & " "
                
                If l = iCount - 1 Then
                    sSQLAGN = sSQLAGN & " ) "
                Else
                    sSQLAGN = sSQLAGN & " or "
                End If
                
                rsKASQLAGN.MoveNext
            Next l
        End If
        rsKASQLAGN.Close
    
    pbrZeit.Value = 200
    If (sKaufdatVon <> "" And sKaufdatBis <> "") Or iAGN <> 0 Or lLieferant <> 0 Then '***mit kassjour
        sSQLKunden = sSQLKunden & ", kassjour "
    End If
    
    If sKredit <> "" Then '***mit kredit
        sSQLKunden = sSQLKunden & ", kredit "
    End If
    
    sSQLKunden = sSQLKunden & " where "
    
'    If sPlzVon <> "" Then
        sSQLKunden = sSQLKunden & " PLZ like '" & sPlzVon & "*' "
'    End If

    If sawm <> "" Then
        sSQLKunden = sSQLKunden & " and awm = '" & sawm & "' "
    End If
    
    If (sKaufdatVon <> "" And sKaufdatBis <> "") Or iAGN <> 0 Or lLieferant <> 0 Then  '***mit kassjour
        sSQLKunden = sSQLKunden & " and "
       
        sSQLKunden = sSQLKunden & " kassjour.kundnr = kunden.kundnr  "
        
    End If
    
    If sKredit <> "" Then '***mit kredit
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " kredit.kundnr = kunden.kundnr  "
    End If
    
    If iGebMonat <> "0" Then '***mit Geburtsmonat
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " Month(Datum1) = "
        sSQLKunden = sSQLKunden & iGebMonat
    End If
    
    If sGeschlecht <> "" Then
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " Geschlecht = "
        sSQLKunden = sSQLKunden & "'" & sGeschlecht & "' "
    End If
    
    If bDS Then
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " DS = True "
    End If
    
    If lKdNumVon <> 0 And lKdnumBis <> 0 Then
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " Kunden.KundNr Between " & lKdNumVon & " and " & lKdnumBis & " "
    End If
    
    If iAGN <> 0 Then '***mit kassjour
        sSQLKunden = sSQLKunden & sSQLAGN
    End If
    
    If lLieferant <> 0 Then '***mit kassjour
        sSQLKunden = sSQLKunden & sSQLLL
    End If
    
    
    
    If sKaufdatVon <> "" And sKaufdatBis <> "" Then '***mit kassjour
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " kassjour.adate Between " & sKaufdatVon & " and " & sKaufdatBis & " "
    End If
    
    If dBowertVon <> 0 And dBowertBis <> 0 Then
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " Bonus between " & dBowertVon & " and " & dBowertBis & ""
    End If

    If sDat1Von <> "" And sDat1Bis <> "" Then
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " datum1  Between " & sDat1Von & " and " & sDat1Bis & " "
    End If
    
    If sDat2Von <> "" And sDat2Bis <> "" Then
        sSQLKunden = sSQLKunden & " and "
        sSQLKunden = sSQLKunden & " datum2 Between " & sDat2Von & " and " & sDat2Bis & " "
    End If
    
        
    '******************************ORTE
        
        sSQL = "Select * From KASQLORT Where KANAME = '" & sdatname & "'"
        Set rsKASQLORT = gdBase.OpenRecordset(sSQL)

        If Not rsKASQLORT.EOF Then
        
            sSQLKunden = sSQLKunden & " and ("
            rsKASQLORT.MoveLast
            
            iCount = rsKASQLORT.RecordCount
            
            rsKASQLORT.MoveFirst
            
            For l = 0 To iCount - 1

            
                If Not IsNull(rsKASQLORT!Ort) Then
                    sOrt = rsKASQLORT!Ort
                Else
                    sOrt = ""
                End If
        
                sSQLKunden = sSQLKunden & " Stadt like '" & sOrt & "*'"
                
                If l = iCount - 1 Then
                    sSQLKunden = sSQLKunden & " ) "
                Else
                    sSQLKunden = sSQLKunden & " or "
                End If
                
                rsKASQLORT.MoveNext
            Next l
        End If
        rsKASQLORT.Close
    '******************************Filiale
        
        sSQL = "Select * From KASQLFIL Where KANAME = '" & sdatname & "'"
        Set rsKASQLFIL = gdBase.OpenRecordset(sSQL)

        If Not rsKASQLFIL.EOF Then
        
            sSQLKunden = sSQLKunden & " and ("
            rsKASQLFIL.MoveLast
            
            iCount = rsKASQLFIL.RecordCount
            
            rsKASQLFIL.MoveFirst
            
            For l = 0 To iCount - 1

            
                If Not IsNull(rsKASQLFIL!FILIALE) Then
                    byFil = rsKASQLFIL!FILIALE
                Else
                    byFil = 0
                End If
        
                sSQLKunden = sSQLKunden & " FilialNr = " & byFil & " "
                
                If l = iCount - 1 Then
                    sSQLKunden = sSQLKunden & " ) "
                Else
                    sSQLKunden = sSQLKunden & " or "
                End If
                
                rsKASQLFIL.MoveNext
            Next l
        End If
        rsKASQLFIL.Close
    '************************Merkmal
        
        sSQL = "Select * From KASQLMK Where KANAME = '" & sdatname & "'"
        Set rsKASQLMK = gdBase.OpenRecordset(sSQL)

        If Not rsKASQLMK.EOF Then
        
            sSQLKunden = sSQLKunden & " and ("
            rsKASQLMK.MoveLast
            
            iCount = rsKASQLMK.RecordCount
            rsKASQLMK.MoveFirst
            
            For l = 0 To iCount - 1

            
                If Not IsNull(rsKASQLMK!MERKMAL) Then
                    sMerkmal = rsKASQLMK!MERKMAL
                Else
                    sMerkmal = ""
                End If
        
                sSQLKunden = sSQLKunden & " Merkmal like '*" & sMerkmal & "*'"
                
                If l = iCount - 1 Then
                    sSQLKunden = sSQLKunden & " ) "
                Else
                    sSQLKunden = sSQLKunden & " or "
                End If
                
                rsKASQLMK.MoveNext
            Next l
        End If
        rsKASQLMK.Close
        
    sSQLKunden = sSQLKunden & " and Kunden.Status <> 'D' "
    sSQLKunden = sSQLKunden & " group by Kunden.KUNDNR "
    sSQLKunden = sSQLKunden & ", Kunden.KUERZEL "
    sSQLKunden = sSQLKunden & ", Kunden.FIRMA "
    sSQLKunden = sSQLKunden & ", Kunden.TITEL "
    sSQLKunden = sSQLKunden & ", Kunden.NAME "
    sSQLKunden = sSQLKunden & ", Kunden.VORNAME "
    sSQLKunden = sSQLKunden & ", Kunden.STRASSE "
    sSQLKunden = sSQLKunden & ", Kunden.PLZ "
    sSQLKunden = sSQLKunden & ", Kunden.STADT "
    sSQLKunden = sSQLKunden & ", Kunden.TEL "
    sSQLKunden = sSQLKunden & ", Kunden.FAXNR "
    sSQLKunden = sSQLKunden & ", Kunden.MERKMAL "
    sSQLKunden = sSQLKunden & ", Kunden.ANREDE "
    sSQLKunden = sSQLKunden & ", Kunden.MERKMAL2 "
    sSQLKunden = sSQLKunden & ", Kunden.FORMATDAT "
    sSQLKunden = sSQLKunden & ", Kunden.RECHNR "
    sSQLKunden = sSQLKunden & ", Kunden.KURZTEXT1 "
    sSQLKunden = sSQLKunden & ", Kunden.KURZTEXT2 "
    sSQLKunden = sSQLKunden & ", Kunden.DATUM1 "
    sSQLKunden = sSQLKunden & ", Kunden.DATUM2 "
    sSQLKunden = sSQLKunden & ", Kunden.UMSLJ "
    sSQLKunden = sSQLKunden & ", Kunden.UMSVJ "
    sSQLKunden = sSQLKunden & ", Kunden.OSUM "
    sSQLKunden = sSQLKunden & ", Kunden.KASSE "
    sSQLKunden = sSQLKunden & ", Kunden.RABATT "
    sSQLKunden = sSQLKunden & ", Kunden.FILIALNR "
    sSQLKunden = sSQLKunden & ", Kunden.GESCHLECHT "
    sSQLKunden = sSQLKunden & ", Kunden.ECIDENT "
    sSQLKunden = sSQLKunden & ", Kunden.GESPERRT "
    sSQLKunden = sSQLKunden & ", Kunden.KUNDKART "
    sSQLKunden = sSQLKunden & ", Kunden.BONUS "
    sSQLKunden = sSQLKunden & ", Kunden.PREISKZ "
    sSQLKunden = sSQLKunden & ", Kunden.Angelegt "
    sSQLKunden = sSQLKunden & ", Kunden.Aender "
    sSQLKunden = sSQLKunden & ", Kunden.Lastdate "
    sSQLKunden = sSQLKunden & ", Kunden.Lasttime "
    sSQLKunden = sSQLKunden & ", Kunden.EMAIL "
    sSQLKunden = sSQLKunden & ", Kunden.MOBILTEL "
    sSQLKunden = sSQLKunden & ", Kunden.awm "
    sSQLKunden = sSQLKunden & ", Kunden.DS "
    
'    MsgBox sSQLKunden
    
    pbrZeit.Value = 250
    gdBase.Execute sSQLKunden, dbFailOnError
    
    


    Dim sSQLKuteilme As String
    
    If dUmsatzVon <> 0 And dUmsatzBis <> 0 Then '*** mit Kassjour
        loeschNEW "Kuteil", gdBase
        
        sSQLKuteilme = " Select   "
        sSQLKuteilme = sSQLKuteilme & " KUTEILME.Knummer"
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUERZEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FIRMA "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TITEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.NAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.VORNAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STRASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PLZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STADT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FAXNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ANREDE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FORMATDAT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RECHNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSLJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSVJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.OSUM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RABATT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FILIALNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESCHLECHT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ECIDENT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESPERRT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUNDKART "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.BONUS "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PREISKZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Angelegt "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Aender "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lastdate "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lasttime "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.EMAIL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.AWM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DS "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MOBILTEL "
        sSQLKuteilme = sSQLKuteilme & ", ' ' as OKredit "
        sSQLKuteilme = sSQLKuteilme & ", '" & sdatname & "' as Datname "
        sSQLKuteilme = sSQLKuteilme & ", '" & sErstelldat & "' as Daterstellung "
        sSQLKuteilme = sSQLKuteilme & ", 0 as Ertrag"

        sSQLKuteilme = sSQLKuteilme & ", sum(kassjour.preis) as Umsatz"
    
        sSQLKuteilme = sSQLKuteilme & " into KuTeil "
        sSQLKuteilme = sSQLKuteilme & " from KUTEILME, kassjour "
        
        sSQLKuteilme = sSQLKuteilme & " where  KUTEILME.knummer = kassjour.kundnr "
        
        If iAGN <> 0 Then '***mit kassjour, AGN
        sSQLKuteilme = sSQLKuteilme & sSQLAGN
        End If
    
        If lLieferant <> 0 Then '***mit kassjour, LINR + LPZ
        sSQLKuteilme = sSQLKuteilme & sSQLLL
        End If
        
        If sKaufdatVon <> "" And sKaufdatBis <> "" Then '***mit kassjour, Kaufdat
        sSQLKuteilme = sSQLKuteilme & " and "
        sSQLKuteilme = sSQLKuteilme & " kassjour.adate Between " & sKaufdatVon & " and " & sKaufdatBis & " "
        End If
        
        sSQLKuteilme = sSQLKuteilme & " group by KUTEILME.Knummer "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUERZEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FIRMA "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TITEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.NAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.VORNAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STRASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PLZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STADT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FAXNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ANREDE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FORMATDAT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RECHNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSLJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSVJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.OSUM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RABATT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FILIALNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESCHLECHT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ECIDENT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESPERRT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUNDKART "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.BONUS "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PREISKZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Angelegt "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Aender "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lastdate "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lasttime "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.EMAIL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MOBILTEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.AWM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DS "
        
        gdBase.Execute sSQLKuteilme, dbFailOnError
        
        loeschNEW "KUTEILME", gdBase
        
        sSQLKuteilme = "Select * into KUTEILME from kuteil where"
        sSQLKuteilme = sSQLKuteilme & " umsatz between " & dUmsatzVon & " and " & dUmsatzBis & ""
        gdBase.Execute sSQLKuteilme, dbFailOnError
    End If
    
    '*****Ertrag anfang
    
    If dErtragVon <> 0 And dErtragBis <> 0 Then '*** mit Kassjour
        loeschNEW "Kuteil", gdBase
        
        sSQLKuteilme = " Select   "
        sSQLKuteilme = sSQLKuteilme & " KUTEILME.Knummer"
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUERZEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FIRMA "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TITEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.NAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.VORNAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STRASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PLZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STADT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FAXNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ANREDE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FORMATDAT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RECHNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSLJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSVJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.OSUM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RABATT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FILIALNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESCHLECHT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ECIDENT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESPERRT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUNDKART "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.BONUS "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PREISKZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Angelegt "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Aender "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lastdate "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lasttime "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.EMAIL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MOBILTEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.AWM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DS "
        sSQLKuteilme = sSQLKuteilme & ", ' ' as OKredit "
        sSQLKuteilme = sSQLKuteilme & ", '" & sdatname & "' as Datname "
        sSQLKuteilme = sSQLKuteilme & ", '" & sErstelldat & "' as Daterstellung "
        sSQLKuteilme = sSQLKuteilme & ", sum((kassjour.preis)-(kassjour.menge * kassjour.ekpr))as Ertrag"
        sSQLKuteilme = sSQLKuteilme & ", 0 as Umsatz"
        sSQLKuteilme = sSQLKuteilme & " into KuTeil "
        sSQLKuteilme = sSQLKuteilme & " from KUTEILME, kassjour "
        
        sSQLKuteilme = sSQLKuteilme & " where  KUTEILME.knummer = kassjour.kundnr "
        
        If iAGN <> 0 Then '***mit kassjour, AGN
        sSQLKuteilme = sSQLKuteilme & sSQLAGN
        End If
    
        If lLieferant <> 0 Then '***mit kassjour, LINR + LPZ
        sSQLKuteilme = sSQLKuteilme & sSQLLL
        End If
        
        If sKaufdatVon <> "" And sKaufdatBis <> "" Then '***mit kassjour, Kaufdat
        sSQLKuteilme = sSQLKuteilme & " and "
        sSQLKuteilme = sSQLKuteilme & " kassjour.adate Between " & sKaufdatVon & " and " & sKaufdatBis & " "
        End If
        
        sSQLKuteilme = sSQLKuteilme & " group by KUTEILME.Knummer "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUERZEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FIRMA "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TITEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.NAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.VORNAME "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STRASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PLZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.STADT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.TEL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FAXNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ANREDE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MERKMAL2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FORMATDAT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RECHNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KURZTEXT2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM1 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DATUM2 "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSLJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.UMSVJ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.OSUM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KASSE "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.RABATT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.FILIALNR "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESCHLECHT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.ECIDENT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.GESPERRT "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.KUNDKART "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.BONUS "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.PREISKZ "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Angelegt "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Aender "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lastdate "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.Lasttime "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.AWM "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.DS "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.EMAIL "
        sSQLKuteilme = sSQLKuteilme & ", KUTEILME.MOBILTEL "
        
        gdBase.Execute sSQLKuteilme, dbFailOnError
        
        loeschNEW "KUTEILME", gdBase
        
        sSQLKuteilme = "Select * into KUTEILME from kuteil where"
        sSQLKuteilme = sSQLKuteilme & " ertrag between " & dErtragVon & " and " & dErtragBis & ""
        
        gdBase.Execute sSQLKuteilme, dbFailOnError
    
    End If
    
    pbrZeit.Value = 300
    
    sSQL = " Delete From KASQL where KAname = 'kein Betreff' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLLL where KAname = 'kein Betreff' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLMK where KAname = 'kein Betreff' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLORT where KAname = 'kein Betreff' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLAGN where KAname = 'kein Betreff' "
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLFIL where KAname = 'kein Betreff' "
    gdBase.Execute sSQL, dbFailOnError
    
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ausführen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub cmdFormat_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    cmdFormat(Index).Enabled = False
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cPfad1      As String
    Dim i           As Integer
    Dim cDatname    As String
    Dim rsrs        As Recordset
    Dim dUmsatzLJ   As Double
    Dim dUmsatzVJ   As Double
    
    cDatname = "KundenA" & Format$(TimeValue(Now), "HH:MM:SS")
    cDatname = SwapStr(cDatname, ":", "")
    cDatname = cDatname & ".xls"
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    If NewTableSuchenDBKombi("KUTEILME", gdBase) Then
    
        Select Case Index
            Case Is = 0     'EXCEL
            
                    
                    loeschNEW "KunExc", gdBase
                    
                    anzeige "normal", "Ausgabe 1/10", lblanzeige
                    
                    gsZSpalte = ""
                    gstab = "KUEX"
                    frmWKL36.Show 1
                    
                    Screen.MousePointer = 11
                    
                    'dannach Tablay auswerten
                    
                    Tabcheck "KUEX"
                    FormatGridOverTablay "KUEX"
                    
                    loeschNEW "Kunden_UMS_lf", gdBase
                    loeschNEW "KASS_TEMPO", gdBase
                    
                    sSQL = "Select * into KASS_TEMPO from kassjour where year(adate)= year(now) "
                    sSQL = sSQL & " and ums_ok = 'J'  and kundnr > 0"
                    gdBase.Execute sSQL, dbFailOnError
                    
                    anzeige "normal", "Ausgabe 2/10", lblanzeige
                    
                    CheckIndex "KASS_TEMPO", "kundnr", "", gdBase
                    
                    anzeige "normal", "Ausgabe 3/10", lblanzeige
                    
                    CheckIndex "KUTEILME", "knummer", "", gdBase
                    
                    anzeige "normal", "Ausgabe 4/10", lblanzeige
                    
                    sSQL = "Select sum(preis) as UMSLJ,kundnr into Kunden_UMS_lf from KASS_TEMPO  "
                    sSQL = sSQL & " where kundnr in (Select knummer from KUTEILME) group by kundnr "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    anzeige "normal", "Ausgabe 5/10", lblanzeige
                    
                    sSQL = "Update KUTEILME set UMSLJ = 0 "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Update KUTEILME k inner join Kunden_UMS_lf u on k.knummer = u.kundnr "
                    sSQL = sSQL & " set k.UMSLJ = u.UMSLJ "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    anzeige "normal", "Ausgabe 6/10", lblanzeige
                    
                    loeschNEW "KASS_TEMPO", gdBase
                    
                    Dim lJahr As Long
    
                    lJahr = Year(Now)
    
                    sSQL = "Select * into KASS_TEMPO from kassjour where year(adate)= " & lJahr & " -1 "
                    sSQL = sSQL & " and ums_ok = 'J'  and kundnr > 0"
                    gdBase.Execute sSQL, dbFailOnError
                    
                    anzeige "normal", "Ausgabe 7/10", lblanzeige
                    
                    CheckIndex "KASS_TEMPO", "kundnr", "", gdBase
                    
                    anzeige "normal", "Ausgabe 8/10", lblanzeige
                    
                    loeschNEW "Kunden_UMS_lf", gdBase
                    
                    sSQL = "Select sum(preis) as UMSVJ,kundnr into Kunden_UMS_lf from KASS_TEMPO  "
                    sSQL = sSQL & " where kundnr in (Select knummer from KUTEILME) group by kundnr "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    anzeige "normal", "Ausgabe 9/10", lblanzeige
                    
                    sSQL = "Update KUTEILME set UMSVJ = 0 "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Update KUTEILME k inner join Kunden_UMS_lf u on k.knummer = u.kundnr "
                    sSQL = sSQL & " set k.UMSVJ = u.UMSVJ "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    
                    
                    
                    
                    
                    
                    
                    

                    
                    sSQL = " Update KUTEILME set Datum1 = 0 where Datum1 = '' "
                    gdBase.Execute sSQL, dbFailOnError
                    sSQL = " Update KUTEILME set Datum1 = 0 where Datum1 is null "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    
                    
                    'Anrede BA
                    SpalteAnfuegenNEW "KUTEILME", "ANREDE_BA", "TEXT(60)", gdBase
                
                    sSQL = "Update KUTEILME set ANREDE_BA = ANREDE"
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Update KUTEILME set ANREDE_BA = 'Herrn' where Ucase(ANREDE_BA) = 'HERR' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Update KUTEILME set ANREDE_BA = 'Herrn' where Ucase(GESCHLECHT) = 'M' "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    
                    
                    
                    
                    
                    anzeige "normal", "", lblanzeige
                
                    If byAnzahlSpalten > 0 Then
                        sSQL = "Select " & sSpaltenbez(0) & " "
                        
                        If byAnzahlSpalten > 1 Then
                            For i = 1 To byAnzahlSpalten - 1
                                sSQL = sSQL & " , " & sSpaltenbez(i) & " "
                            Next i
                        End If
                    Else
                        Exit Sub
                    End If
                    sSQL = sSQL & " into KunExc from KUTEILME"
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "alter table KUTEILME drop column ANREDE_BA"
                    gdBase.Execute sSQL, dbFailOnError
                    
                    Screen.MousePointer = 0
                    
                    If bDat Then
                    
                        If gsKUPFAD <> "" Then
                            cdatei = gsKUPFAD
                            Kill gsKUPFAD
                        Else
                            cdatei = cPfad1 & "BOX\" & cDatname
                            cPfad = cPfad1 & "BOX"
                        End If
                        
                        sSQL = "Select * "
                        sSQL = sSQL & " into KunExc IN '" & cdatei & "' 'Excel 8.0;' from KunExc "
                        gdBase.Execute sSQL, dbFailOnError
                        
                        If gsKUPFAD <> "" Then
                            MsgBox "Diese Datei ist unter (" & gsKUPFAD & ") abgespeichert", vbInformation, "Winkiss Information:"
                        Else
                            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & cDatname & " abgespeichert", vbInformation, "Winkiss Information:"
                        End If
                    ElseIf bEmail Then
                    
                        cdatei = cPfad1 & "BOX\" & cDatname
                        cPfad = cPfad1 & "BOX"
                        
                        sSQL = "Select * into KunExc IN '" & cdatei & "' 'Excel 8.0;' from KunExc "
                        gdBase.Execute sSQL, dbFailOnError
                    
                        bExcel = True
        
                        fraEmail.Visible = True
                        Text1(0).SetFocus
        
                    End If
            Case Is = 1     'Word bzw RTF
                If bDat Then
                    cr2.ReportFileName = cPfad1 & "aWKLavg.rpt"
                    cr2.PrintFileName = cPfad1 & "BOX\Kundenliste.doc"
                    cr2.PrintFileType = crptRTF
                    cr2.Destination = 2
                    cr2.Action = 1
                    
                    MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: Kundenliste.doc abgespeichert", vbInformation, "Winkiss Information:"
    
                ElseIf bEmail Then
                    bWord = True
                    fraEmail.Visible = True
                    Text1(0).SetFocus
                   
                End If
                
            Case Is = 3     'Steuerdatei einfach
                
                loeschNEW "stdatei", gdBase
            
                sSQL = "Select KNUMMER, KUERZEL, FIRMA, TITEL, NAME, VORNAME, STRASSE, PLZ, STADT, TEL, FAXNR "
                sSQL = sSQL & ", ANREDE,Kurztext1,Datum1,Geschlecht into Stdatei from KUTEILME"
                gdBase.Execute sSQL, dbFailOnError
                
                If bDat Then
                    cdatei = cPfad1 & "BOX\StDatei.dbf"
                    cPfad = cPfad1 & "BOX"
                    Kill cdatei
                    
                    If NewTableSuchenDBKombi("StDatei", gdBase) Then
                        sSQL = "Select * into StDatei IN '" & cPfad & "' 'dbase IV;' from StDatei "
                        gdBase.Execute sSQL, dbFailOnError
    
                        Screen.MousePointer = 0
                        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: StDatei.dbf abgespeichert", vbInformation, "Winkiss Information:"
                    End If
                End If
                
            Case Is = 2     'Steuerdatei erweitert
            
                loeschNEW "stdater", gdBase
                
                sSQL = "Select * into Stdater from KUTEILME"
                gdBase.Execute sSQL, dbFailOnError
                
                If bDat Then
                    cdatei = cPfad1 & "BOX\StDater.dbf"
                    cPfad = cPfad1 & "BOX"
                    Kill cdatei
                
                    If NewTableSuchenDBKombi("StDater", gdBase) Then
                        sSQL = "Select * into StDater IN '" & cPfad & "' 'dbase IV;' from StDater "
                        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
                        MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: StDater.dbf abgespeichert", vbInformation, "Winkiss Information:"
                    End If
                End If
        End Select
    
    Else
        anzeige "rot", "Bitte erst Kunden ermitteln - dann die Ausgabeart bestimmen!", lblanzeige
    End If
    
    cmdFormat(Index).Enabled = True

    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 75 Then
        Resume Next
    ElseIf err.Number = 20530 Or err.Number = 3051 Then
        Screen.MousePointer = 0
        MsgBox "Sie haben keine Diskette eingelegt", vbInformation, "Winkiss Hinweis"
    ElseIf err.Number = 20999 Then
        Screen.MousePointer = 0
        MsgBox "Bitte nutzen Sie ein anderes Ausgabeformat! Die Ausgabe in diesem Format ist nicht möglich. ", vbInformation, "Winkiss Hinweis"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "cmdFormat_Click"
        Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten. " & Index
        
        Fehlermeldung1
    End If
End Sub
Private Sub cmdSUeber_Click()
    On Error GoTo LOKAL_ERROR
    Dim sHaupt As String
    Dim sSQL As String
    Dim sPfad As String
    
    sPfad = gcDBPfad
    If Right(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    
    fraSerienB.Visible = False
    
    sHaupt = txtSerienBHaupt.Text
    
    loeschNEW "Haupt", gdBase
    
    sSQL = "create table haupt ("
    sSQL = sSQL & " texthaupt memo )"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into haupt"
    sSQL = sSQL & "(texthaupt) "
    sSQL = sSQL & "values ("
    sSQL = sSQL & "'" & sHaupt & "' "
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    Pause (1)
    reportbildschirm "kaser", "aWKLavh"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdSUeber_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim lrow As Long
Dim cFarbkenn As String
Dim iRet As Integer
Dim ctmp As String

Select Case Index

    Case 0 'kundendaten
        
        lrow = Val(MSHFLEX1.Row)
        If lrow > 0 Then
            MSHFLEX1.Row = lrow
            MSHFLEX1.Col = 1
            gcKundenNr = MSHFLEX1.Text
            iKasse = 2
            frmWKL13.Show 1
        End If
    Case 1
        Screen.MousePointer = 0
                
        gsBackcolor = Label1(16).BackColor
        gsForecolor = Label1(16).ForeColor
        gsKundenFarbe = Label1(16).Tag
        
        frmWKL65.Show 1
        
        Label1(16).BackColor = gsBackcolor
        Label1(16).ForeColor = glS1
        Label1(16).Tag = gsKundenFarbe
        If gsKundenFarbe <> "" Then
            Label1(16).Caption = "Farbauswahl"
        Else
            Label1(16).Caption = "alle Farben"
        End If
    Case 2 'historie
        lrow = Val(MSHFLEX1.Row)
        If lrow > 0 Then
            MSHFLEX1.Row = lrow
            MSHFLEX1.Col = 1
            gckundnr = MSHFLEX1.Text
            
            gckundnr = Trim$(gckundnr)
            gsARTNR = ""
            
            frmWKL74.Show 1
        End If
    Case 3
        If Command1(3).Caption = "alle zurücksetzen" Then
        
            SchalteKunden (2)
            Command1(3).Caption = "alle auswählen"
        ElseIf Command1(3).Caption = "alle auswählen" Then
            SchalteKunden (3)
            Command1(3).Caption = "alle zurücksetzen"
        End If
    Case 5
        Screen.MousePointer = 0
        
        gsBackcolor = Label4(32).BackColor
        gsForecolor = Label4(32).ForeColor
        gsKundenFarbe = Label4(32).Tag
        
        frmWKL65.Show 1
        
        Label4(32).BackColor = gsBackcolor
        Label4(32).ForeColor = gsForecolor
        Label4(32).Tag = gsKundenFarbe
        If gsKundenFarbe <> "" Then
            Label4(32).Caption = "Farbauswahl"
        Else
            Label4(32).Caption = "alle Farben"
        End If
        
    Case 6
        ctmp = Trim$(Label4(32).Tag)
        If ctmp <> "" Then
            cFarbkenn = ermFarbeKU(ctmp)
        Else
            cFarbkenn = "alle Farben"
            SchalteKunden (2)
            Exit Sub
            ctmp = "0"
        End If
        
        If cFarbkenn = "" Then cFarbkenn = "ohne Kennzeichen"
        
        iRet = MsgBox("Möchten Sie jetzt alle Kunden aus der Tabelle mit dem Farbkennzeichen '" & cFarbkenn & "' zurücksetzen?", vbYesNo + vbQuestion + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            Screen.MousePointer = 11
            SchalteKunden (4)
            Screen.MousePointer = 0
            
        End If
        
End Select
            
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchalteKunden(iSchaltung As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow    As Long
    Dim lRows   As Long
    Dim lcol    As Long
    Dim ctmp    As String
    Dim cAWM    As String
    Dim sKUNDNR As String
    
    If iSchaltung = 3 Then
        lAusgewählt = 0
    End If
    
    If iSchaltung = 2 Then
        lAusgewählt = 0
    End If
    
    
    
    lRows = MSHFLEX1.Rows
    lRows = lRows - 1
    lcol = 0
    MSHFLEX1.Redraw = False
    For lrow = 1 To lRows
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        If iSchaltung = 2 Then
            MSHFLEX1.Text = ""
        End If
        If iSchaltung = 4 Then
        
            'ja aber hat der kunden bestimmte farbe
            
           
            anzeige "normal", lrow & "...", lblanzeige
                
            
            ctmp = Trim$(Label4(32).Tag)
            If ctmp = "" Then ctmp = "0"
            
            MSHFLEX1.Col = 1
            sKUNDNR = MSHFLEX1.Text
            
            cAWM = ""
            If sKUNDNR <> "" Then
                cAWM = WhatIsAwmKU(sKUNDNR)
            End If
            
            If cAWM = ctmp Then
                MSHFLEX1.Row = lrow
                MSHFLEX1.Col = lcol
                MSHFLEX1.Text = ""
                lAusgewählt = lAusgewählt - 1
            End If
        End If
        
        If iSchaltung = 3 Then
            MSHFLEX1.Text = "X"
            lAusgewählt = lAusgewählt + 1
        End If
    Next lrow
    
    MSHFLEX1.Redraw = True
    
    If lAusgewählt > 1 Then
        anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label18
    ElseIf lAusgewählt = 1 Then
        anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label18
    Else
        anzeige "normal", "", Label18
    End If
    
    With MSHFLEX1
        .Row = 1
        .Col = 0
        .SetFocus
    End With
    
Exit Sub
LOKAL_ERROR:
    

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchalteKunden"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub

Private Sub Command2_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            frmWKL124.Show 1
        Case 1
            frmWKL125.Show 1
        Case 2
            frmWKL162.Show 1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    Dim lcount As Long
    
    Select Case Index
        Case Is = 0        ' Kalender
            txtDat2(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            txtDat2(1).SetFocus
            
        Case Is = 1        ' Kalender
            txtDat2(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
        Case Is = 30        ' Kalender
            txtKauf(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            txtKauf(1).SetFocus
            
        Case Is = 31        ' Kalender
            txtKauf(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
        Case Is = 20        ' Kalender
            txtDat1(0).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            txtDat1(1).SetFocus
        Case Is = 21        ' Kalender
            txtDat1(1).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
            'fertig
        Case 9
            txtlief_KeyUp vbKeyF2, 0
        Case 2
        
            gF2Prompt.cFeld = ""
            gF2Prompt.cWert = ""
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            
            gF2Prompt.bMultiple = True
            gF2Prompt.cFeld = "AGN"
            If gF2Prompt.cFeld <> "" Then
                frmWK00a.Show 1
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                End If
            End If
            
            lstAGN.Visible = False
            lstAGN.Clear
            For lcount = 0 To 100
                If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                    lstAGN.Visible = True
'                    Text1(Index).Text = ""
                    
                    If gF2Prompt.cArray(lcount) <> "" Then
                        lstAGN.AddItem Mid(gF2Prompt.cArray(lcount), 1, InStr(1, gF2Prompt.cArray(lcount), " ")) & " "
'                        lstAGN.AddItem gF2Prompt.cArray(lCount)
                    End If
                Else
                    If gF2Prompt.cArray(lcount) <> "" Then
'                        lstAGN.AddItem gF2Prompt.cArray(lCount)
                        lstAGN.AddItem Mid(gF2Prompt.cArray(lcount), 1, InStr(1, gF2Prompt.cArray(lcount), " ")) & " "
'                        Text1(Index).Text = Mid(gF2Prompt.cArray(lcount), 1, InStr(1, gF2Prompt.cArray(lcount), " ")) & " "
                    End If
                    
                End If
            Next lcount
        
        
        
        
        
        
        
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WKLavPositionieren()
    On Error GoTo LOKAL_ERROR
    
    With MSHFLEX1
        .Height = 5655
        .Left = 480
        .Top = 960
        .Width = 8175
    End With
    
    With fraZuErstellen
        .Top = 1080
        .Left = 480
        .Height = 5655
        .Width = 10815
    End With
    
    With fraAusgabe
        .Top = 3840
        .Left = 3120
        .Height = 2535
        .Width = 6855
    End With
    
    With fraListen
        .Top = 120
        .Left = 2520
        .Height = 2295
        .Width = 2175
    End With
    
    With fraEtiketten
        .Top = 200
        .Left = 2520
        .Height = 2175
        .Width = 2175
    End With
    
    With fraEmail
        .Top = 1200
        .Left = 3120
        .Height = 2415
        .Width = 6855
    End With
    
    With fraSerienB
        .Top = 1200
        .Left = 3120
        .Height = 2415
        .Width = 6855
    End With
    
    With fraFormat
        .Top = 120
        .Left = 4680
        .Height = 1815
        .Width = 2055
    End With
    
    With fraSort
        .Top = 120
        .Left = 4680
        .Height = 1815
        .Width = 2055
    End With
    
    With fraExport
        .Top = 960
        .Left = 2520
        .Height = 1455
        .Width = 2175
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKLavPositionieren"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    WKLavPositionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift

    bKauf = False
    bDat1 = False
    bDat2 = False
    bVorhanden = False
    bAender = False
    bNotAll = False
    bClickAusgabe = False
    bEmail = False
'    bDis = False
    bDat = False
    bExcel = False
    bWord = False
    
    sdateiname = ""
    sErstelldatum = ""
    
    fraZuErstellen.Caption = "Neue Zusammenstellung"
    
    
    anzeige "normal", "Geben Sie in diesem Formular ihre Suchkriterien ein!", lblanzeige
    
    
    Tabellenvorhanden
    
    sSQL = " Delete From KASQL where KAname is null"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLLL where KAname is null"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLMK where KAname is null"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLORT where KAname is null"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLAGN where KAname is null"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = " Delete From KASQLFIL where KAname is null"
    gdBase.Execute sSQL, dbFailOnError
    
    füllecboAgn
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    LogtoEnd Me
    loeschNEW "KUTEILME", gdBase 'Kundenanalyse
    loeschNEW "Kuteil", gdBase
    loeschNEW "KUTTEN", gdBase
    loeschNEW "Kunden_UMS_lf", gdBase
    loeschNEW "KASS_TEMPO", gdBase

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_DblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 16 Then
        Label1(16).Caption = "alle Farben"
        Label1(16).Tag = ""
        Label1(16).BackColor = glH1 'Label1(14).BackColor
        Label1(16).ForeColor = glS1
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub lstAGN_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = 46    'Del
            If Not lstAGN.ListIndex = -1 Then
                lstAGN.RemoveItem (lstAGN.ListIndex)
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstAGN_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstdatnames_Click()
    On Error GoTo LOKAL_ERROR

    MSHFLEX1.Visible = False
    
    anzeige "normal", "", lblanzeige
    
    sdateiname = Right(lstdatnames.list(lstdatnames.ListIndex), Len(lstdatnames.list(lstdatnames.ListIndex)) - 11)
    sErstelldatum = Left(lstdatnames.list(lstdatnames.ListIndex), InStr(lstdatnames.list(lstdatnames.ListIndex), " "))
    MousePointer = vbHourglass
    
    ZusammenstellunginMaskezeigen Trim(sdateiname)
    fraZuErstellen.Caption = "Zusammenstellung vom " & sErstelldatum & "      Name der Zusammenstellung: " & sdateiname
    bVorhanden = True
    
    MousePointer = vbDefault
    fraZuErstellen.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstdatnames_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstdatnames_DblClick()
    On Error GoTo LOKAL_ERROR
    
    Dim sdatname As String
    
    If Not lstdatnames.ListIndex = -1 Then
        sdatname = Right(lstdatnames.list(lstdatnames.ListIndex), Len(lstdatnames.list(lstdatnames.ListIndex)) - 11)
        sErstelldatum = Left(lstdatnames.list(lstdatnames.ListIndex), InStr(lstdatnames.list(lstdatnames.ListIndex), " "))
        
        MousePointer = vbHourglass
    
        anzeige "normal", "Die Kundendaten werden ermittelt...", lblanzeige
        
        ausführen Trim(sdatname), sErstelldatum
        Zusammenstellunganzeigen
        
        MousePointer = vbDefault
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstdatnames_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZusammenstellunginMaskezeigen(sdatname As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    Dim rsKASQL         As Recordset
    Dim rsKASQLLL       As Recordset
    Dim rsKASQLMK       As Recordset
    Dim rsKASQLAGN      As Recordset
    Dim rsKASQLORT      As Recordset
    Dim rsKASQLFIL      As Recordset
    Dim rsKd            As Recordset
    
    
    Dim sSQLLL          As String
    
    Dim sGeschlecht     As String
    Dim lKdNumVon       As Long
    Dim lKdnumBis       As Long
    Dim sPlzVon         As String
    Dim sawm            As String
    Dim sKaufdatVon     As String
    Dim sKaufdatBis     As String
    Dim dBowertVon      As Double
    Dim dBowertBis      As Double
    Dim dErtragVon      As Double
    Dim dErtragBis      As Double
    Dim dUmsatzVon      As Double
    Dim dUmsatzBis      As Double
    Dim sDat1Von        As String
    Dim sDat1Bis        As String
    Dim sDat2Von        As String
    Dim sDat2Bis        As String
    Dim sKredit         As String
    Dim sGebmonat       As String
    
    
    Dim lKaufdatVon     As Long
    Dim lKaufdatBis     As Long
    Dim lDat            As Long
    
    Dim iAGN            As Integer
    Dim sOrt            As String
    Dim sMerkmal        As String
    Dim lLieferant      As Long
    Dim sLinie          As String
    Dim byFil           As Byte
    Dim bDS             As Boolean
    
    sSQL = "Select * from KASQL "
    sSQL = sSQL & " where KAname = '" & sdatname & "'"
    
    Set rsKASQL = gdBase.OpenRecordset(sSQL)
        
    If Not rsKASQL.EOF Then
        rsKASQL.MoveFirst
        
        If Not IsNull(rsKASQL!DS) Then
            If rsKASQL!DS = True Then
                bDS = True
            Else
                bDS = False
            End If
        End If
    
        If Not IsNull(rsKASQL!geschlecht) Then
            sGeschlecht = rsKASQL!geschlecht
        Else
            sGeschlecht = ""
        End If
        
        If Not IsNull(rsKASQL!KREDIT) Then
            sKredit = rsKASQL!KREDIT
        Else
            sKredit = ""
        End If
        
        If Not IsNull(rsKASQL!Gebmonat) Then
            If rsKASQL!Gebmonat < 1 Or rsKASQL!Gebmonat > 12 Then
                sGebmonat = ""
            Else
                sGebmonat = MonthName(rsKASQL!Gebmonat)
            End If
        Else
            sGebmonat = ""
        End If
        
        If Not IsNull(rsKASQL!KdNumVon) Then
            lKdNumVon = rsKASQL!KdNumVon
        Else
            lKdNumVon = "0"
        End If
        
        If Not IsNull(rsKASQL!KdnumBis) Then
            lKdnumBis = rsKASQL!KdnumBis
        Else
            lKdnumBis = 0
        End If
        
        If lKdnumBis = 0 And lKdNumVon <> 0 Then
            lKdnumBis = rsKASQL!KdNumVon
        End If
        
        If Not IsNull(rsKASQL!AWM) Then
            sawm = rsKASQL!AWM
        Else
            sawm = ""
        End If
        
        If Not IsNull(rsKASQL!PlzVon) Then
            sPlzVon = rsKASQL!PlzVon
        Else
            sPlzVon = ""
        End If
        
        If rsKASQL!KaufdatVon <> "" Then
'            If Not IsNull(rsKASQL!KaufdatVon) Then
            sKaufdatVon = rsKASQL!KaufdatVon
            lKaufdatVon = DateValue(sKaufdatVon)
            sKaufdatVon = CLng(lKaufdatVon)
        Else
            sKaufdatVon = ""
        End If
        
        If rsKASQL!KaufdatBis <> "" Then
'            If Not IsNull(rsKASQL!KaufdatBis) Then
            sKaufdatBis = rsKASQL!KaufdatBis
            lKaufdatBis = DateValue(sKaufdatBis)
            sKaufdatBis = CLng(lKaufdatBis)
        Else
            sKaufdatBis = ""
        End If
            
        If sKaufdatBis = "" Then
            lKaufdatBis = DateValue(Now)
            sKaufdatBis = CLng(lKaufdatBis)
        End If
        
        If Not IsNull(rsKASQL!BowertVon) Then
            dBowertVon = Format$(rsKASQL!BowertVon, "######0")
        Else
            dBowertVon = 0
        End If
        
        If Not IsNull(rsKASQL!BowertBis) Then
            dBowertBis = Format$(rsKASQL!BowertBis, "######0")
        Else
            dBowertBis = 0
        End If
        
            If dBowertBis = 0 Then
                sSQL = "select Max(Bonus)as MoM from Kunden"
                Set rsKd = gdBase.OpenRecordset(sSQL)
                
                If Not rsKd.EOF Then
                    rsKd.MoveFirst
            
                    If Not IsNull(rsKd!MoM) Then
                        dBowertBis = Format$(rsKd!MoM, "######0")
                    Else
                        dBowertBis = 0
                    End If
                End If
                rsKd.Close
            End If
        
        
        If Not IsNull(rsKASQL!ErtragVon) Then
            dErtragVon = Format$(rsKASQL!ErtragVon, "######0")
        Else
            dErtragVon = 0
        End If
        
        If Not IsNull(rsKASQL!ErtragBis) Then
            dErtragBis = Format$(rsKASQL!ErtragBis, "######0")
        Else
            dErtragBis = 0
        End If
        
        If Not IsNull(rsKASQL!UmsatzVon) Then
            dUmsatzVon = Format$(rsKASQL!UmsatzVon, "######0")
        Else
            dUmsatzVon = 0
        End If
        
        If Not IsNull(rsKASQL!UmsatzBis) Then
            dUmsatzBis = Format$(rsKASQL!UmsatzBis, "######0")
        Else
            dUmsatzBis = 0
        End If
        
        If dUmsatzBis = 0 And dUmsatzVon <> 0 Then
            dUmsatzBis = Format$(rsKASQL!UmsatzVon, "######0")
        End If
        
        If rsKASQL!dat1Von <> "" Then
'            If Not IsNull(rsKASQL!Dat1Von) Then
            sDat1Von = rsKASQL!dat1Von
            lDat = DateValue(sDat1Von)
            sDat1Von = CLng(lDat)
        Else
            sDat1Von = ""
        End If
        
        If rsKASQL!dat1Bis <> "" Then
'            If Not IsNull(rsKASQL!Dat1Bis) Then
            sDat1Bis = rsKASQL!dat1Bis
            lDat = DateValue(sDat1Bis)
            sDat1Bis = CLng(lDat)
        Else
            sDat1Bis = ""
        End If
        
        If sDat1Bis = "" Then
            lDat = DateValue(Now)
            sDat1Bis = CLng(lDat)
        End If
        
        If rsKASQL!dat2Von <> "" Then
'            If Not IsNull(rsKASQL!dat2Von) Then
            sDat2Von = rsKASQL!dat2Von
            lDat = DateValue(sDat2Von)
            sDat2Von = CLng(lDat)
        Else
            sDat2Von = ""
        End If
        
        If rsKASQL!dat2Bis <> "" Then
'            If Not IsNull(rsKASQL!Dat2Bis) Then
            sDat2Bis = rsKASQL!dat2Bis
            lDat = DateValue(sDat2Bis)
            sDat2Bis = CLng(lDat)
        Else
            sDat2Bis = ""
        End If
        
        If sDat2Bis = "" Then
            lDat = DateValue(Now)
            sDat2Bis = CLng(lDat)
        End If
    End If
    
    
    '***Geschlecht
    
    checkweibl.Value = 0
    checkmannl.Value = 0
    
    If sGeschlecht = "W" Then
        checkweibl.Value = 1
    ElseIf sGeschlecht = "M" Then
        checkmannl.Value = 1
    Else
        checkweibl.Value = 0
        checkmannl.Value = 0
    End If
    
    'ds
    
    If bDS = True Then
        chkDS.Value = vbChecked
    Else
        chkDS.Value = vbUnchecked
    End If
    
    '***Kredit
    
    checkOKr.Value = 0
    
    
    If sKredit = "J" Then
        checkOKr.Value = 1
    Else
        checkOKr.Value = 0
    End If
    
   
    
     '***Kundennummer
     
    txtKdNrVon.Text = ""
    txtKdNrBis.Text = ""
    
    If lKdNumVon <> 0 Then
        txtKdNrVon.Text = lKdNumVon
        txtKdNrBis.Text = lKdnumBis
    Else
        txtKdNrVon.Text = ""
        txtKdNrBis.Text = ""
    End If
    
    
    '***Gebmonat
    
    cboGebMonat.Text = ""
    
    If sGebmonat <> "" Then
        cboGebMonat.Text = sGebmonat
    End If
    
    '***PLZ
    
    txtPlzVon.Text = ""
    
    If sPlzVon <> "" Then
        txtPlzVon.Text = sPlzVon
    End If
    
    Label1(16).Caption = "alle Farben"
    Label1(16).Tag = ""
    Label1(16).BackColor = glH1
    Label1(16).ForeColor = Label1(14).ForeColor
    
    If sawm <> "" Then
    
        Label1(16).BackColor = glfarbe(sawm)
        Label1(16).ForeColor = Label1(14).ForeColor



        Label1(16).Tag = sawm
        If sawm <> "" Then
            Label1(16).Caption = "Farbauswahl"
        Else
            Label1(16).Caption = "alle Farben"
        End If
    End If
    
    '***Kaufdatum
    
    txtKauf(0).Text = ""
    txtKauf(1).Text = ""
    
    If sKaufdatVon <> "" Then
        txtKauf(0).Text = rsKASQL!KaufdatVon
        
        
        If Not IsNull(rsKASQL!KaufdatBis) Then
            txtKauf(1).Text = rsKASQL!KaufdatBis
        Else
            txtKauf(1).Text = DateValue(Now)
        End If
        
    Else
        txtKauf(0).Text = ""
        txtKauf(1).Text = ""
    End If
    
    '***Ertrag
    
    txtErtragVon.Text = ""
    txtErtragBis.Text = ""
    
    If dErtragVon <> 0 Then
        txtErtragVon.Text = dErtragVon
        txtErtragBis.Text = dErtragBis
    Else
        txtErtragVon.Text = ""
        txtErtragBis.Text = ""
    End If
    
    '***Bonus
    
    txtBowertVon.Text = ""
    txtBowertBis.Text = ""
    
    If dBowertVon <> 0 Then
        txtBowertVon.Text = dBowertVon
        txtBowertBis.Text = dBowertBis
    Else
        txtBowertVon.Text = ""
        txtBowertBis.Text = ""
    End If
    
    '***Umsatz
    
    txtUmsatzVon.Text = ""
    txtUmsatzBis.Text = ""
    
    If dUmsatzVon <> 0 Then
        txtUmsatzVon.Text = dUmsatzVon
        txtUmsatzBis.Text = dUmsatzBis
    Else
        txtUmsatzVon.Text = ""
        txtUmsatzBis.Text = ""
    End If
    
    '***Datum1
    
    txtDat1(0).Text = ""
    txtDat1(1).Text = ""

    If sDat1Von <> "" Then
        txtDat1(0).Text = rsKASQL!dat1Von
        
        If Not IsNull(rsKASQL!dat1Bis) Then
            txtDat1(1).Text = rsKASQL!dat1Bis
        Else
            txtDat1(1).Text = DateValue(Now)
        End If
        
    Else
        txtDat1(0).Text = ""
        txtDat1(1).Text = ""
    End If
    
    '***Datum2
    
    txtDat2(0).Text = ""
    txtDat2(1).Text = ""

    If sDat2Von <> "" Then
        txtDat2(0).Text = rsKASQL!dat2Von
        If Not IsNull(rsKASQL!dat2Bis) Then
            txtDat2(1).Text = rsKASQL!dat2Bis
        Else
            txtDat2(1).Text = DateValue(Now)
        End If
    Else
        txtDat2(0).Text = ""
        txtDat2(1).Text = ""
    End If
    
    '***Orte
    
    lstOrt.Clear
    
    sSQL = "Select * From KASQLORT Where KANAME = '" & sdatname & "'"
    Set rsKASQLORT = gdBase.OpenRecordset(sSQL)

    If Not rsKASQLORT.EOF Then
        rsKASQLORT.MoveFirst
        Do While Not rsKASQLORT.EOF
            If Not IsNull(rsKASQLORT!Ort) Then
                sOrt = rsKASQLORT!Ort
                lstOrt.AddItem (sOrt)
            Else
                sOrt = ""
            End If
            
        rsKASQLORT.MoveNext
        Loop
    End If
    rsKASQLORT.Close
    
    '***Merkmale
    
    lstMerkmal.Clear
    
    sSQL = "Select * From KASQLMK Where KANAME = '" & sdatname & "'"
    Set rsKASQLMK = gdBase.OpenRecordset(sSQL)

    If Not rsKASQLMK.EOF Then
        rsKASQLMK.MoveFirst
        Do While Not rsKASQLMK.EOF
            If Not IsNull(rsKASQLMK!MERKMAL) Then
                sMerkmal = rsKASQLMK!MERKMAL
                lstMerkmal.AddItem (sMerkmal)
            Else
                sMerkmal = ""
            End If
            
        rsKASQLMK.MoveNext
        Loop
    End If
    rsKASQLMK.Close
    '***AGN
    
    lstAGN.Clear
    
    sSQL = "Select * From KASQLAGN Where KANAME = '" & sdatname & "'"
    Set rsKASQLAGN = gdBase.OpenRecordset(sSQL)

    If Not rsKASQLAGN.EOF Then
        rsKASQLAGN.MoveFirst
        Do While Not rsKASQLAGN.EOF
            If Not IsNull(rsKASQLAGN!AGN) Then
                iAGN = rsKASQLAGN!AGN
                lstAGN.AddItem (iAGN)
            Else
                iAGN = 0
            End If
            
        rsKASQLAGN.MoveNext
        Loop
    End If
    rsKASQLAGN.Close
    '***Filiale
    
    lstFil.Clear
    
    sSQL = "Select * From KASQLFIL Where KANAME = '" & sdatname & "'"
    Set rsKASQLFIL = gdBase.OpenRecordset(sSQL)

    If Not rsKASQLFIL.EOF Then
        rsKASQLFIL.MoveFirst
        Do While Not rsKASQLFIL.EOF
            If Not IsNull(rsKASQLFIL!FILIALE) Then
                byFil = rsKASQLFIL!FILIALE
                lstFil.AddItem (byFil)
            Else
                byFil = 0
            End If
            
        rsKASQLFIL.MoveNext
        Loop
    End If
    rsKASQLFIL.Close
    '***Lieferant Linie
    
    lstLL.Clear
    
    sSQL = "Select * From KASQLLL Where KANAME = '" & sdatname & "'"
    Set rsKASQLLL = gdBase.OpenRecordset(sSQL)

    If Not rsKASQLLL.EOF Then
        rsKASQLLL.MoveFirst
        Do While Not rsKASQLLL.EOF
            If Not IsNull(rsKASQLLL!Lieferant) Then
                lLieferant = rsKASQLLL!Lieferant
            Else
                lLieferant = 0
            End If
            
            
            If Not IsNull(rsKASQLLL!Linie) Then
                sLinie = rsKASQLLL!Linie
            Else
                sLinie = ""
            End If
            
            If lLieferant <> 0 Then
                If sLinie <> "" Then
                    lstLL.AddItem (lLieferant & "   " & sLinie)
                Else
                    lstLL.AddItem (lLieferant)
                End If
            End If
        rsKASQLLL.MoveNext
        Loop
    End If
    rsKASQLLL.Close
    rsKASQL.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZusammenstellunginMaskezeigen"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstdatnames_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sdatname As String
    Dim sSQL As String
    
    
    Select Case KeyCode
        Case Is = 46    'Del
            If Not lstdatnames.ListIndex = -1 Then
                sdatname = Right(lstdatnames.list(lstdatnames.ListIndex), Len(lstdatnames.list(lstdatnames.ListIndex)) - 11)
                lstdatnames.RemoveItem (lstdatnames.ListIndex)
                
                sSQL = " Delete From KASQL where KAname = '" & sdatname & "' "
                gdBase.Execute sSQL, dbFailOnError
                sSQL = " Delete From KASQLLL where KAname = '" & sdatname & "' "
                gdBase.Execute sSQL, dbFailOnError
                sSQL = " Delete From KASQLMK where KAname = '" & sdatname & "' "
                gdBase.Execute sSQL, dbFailOnError
                sSQL = " Delete From KASQLORT where KAname = '" & sdatname & "' "
                gdBase.Execute sSQL, dbFailOnError
                sSQL = " Delete From KASQLAGN where KAname = '" & sdatname & "' "
                gdBase.Execute sSQL, dbFailOnError
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstdatnames_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstFil_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstFil_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstFil_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = 46    'Del
            If Not lstFil.ListIndex = -1 Then
                lstFil.RemoveItem (lstFil.ListIndex)
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstFil_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstLL_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstLL_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstLL_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = 46    'Del
            If Not lstLL.ListIndex = -1 Then
                lstLL.RemoveItem (lstLL.ListIndex)
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstLL_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstMerkmal_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstMerkmal_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstMerkmal_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case KeyCode
        Case Is = 46    'Del
            If Not lstMerkmal.ListIndex = -1 Then
                lstMerkmal.RemoveItem (lstMerkmal.ListIndex)
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstMerkmal_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstOrt_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstOrt_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lstOrt_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case KeyCode
        Case Is = 46    'Del
            If Not lstOrt.ListIndex = -1 Then
                lstOrt.RemoveItem (lstOrt.ListIndex)
            End If
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lstOrt_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_DblClick()
 On Error GoTo LOKAL_ERROR
    Dim lcol As Long
    Dim sSQL As String
    Dim rs As Recordset
    Dim sSortKrit As String
    
    If MSHFLEX1.Row = 1 Then
    
        lcol = MSHFLEX1.Col
        Select Case lcol
            Case Is = 1
            sSortKrit = " order by  Knummer"
            Case Is = 2
            sSortKrit = " order by  Vorname"
            Case Is = 3
            sSortKrit = " order by  Name"
            Case Is = 4
            sSortKrit = " order by  Strasse"
            Case Is = 5
            sSortKrit = " order by  Plz"
            Case Is = 6
            sSortKrit = " order by  stadt"
            Case Is = 7
            sSortKrit = " order by  datum1"
        End Select
        loeschNEW "Kutte", gdBase
        
        
        sSQL = "select * into kutte from KUTEILME " & sSortKrit
        
        If byteSortReihen = 1 Then
            If Trim(sSortKrit) <> "" Then
                sSQL = sSQL & " desc"
            End If
            byteSortReihen = 2
            MSHFLEX1.Col = lcol
            MSHFLEX1.sOrt = 1
        ElseIf byteSortReihen = 2 Then
            If Trim(sSortKrit) <> "" Then
                sSQL = sSQL & " asc"
            End If
            byteSortReihen = 1
            MSHFLEX1.Col = lcol
            MSHFLEX1.sOrt = 2
        End If
        
        gdBase.Execute sSQL
        
        loeschNEW "KUTEILME", gdBase
        sSQL = "select * into KUTEILME from KUTTE "
        gdBase.Execute sSQL
        loeschNEW "Kutte", gdBase
    Else
    
        MSHFLEX1.Col = 0
        If MSHFLEX1.Text = "X" Then
            MSHFLEX1.Text = ""
            lAusgewählt = lAusgewählt - 1
        Else
            MSHFLEX1.Text = "X"
            lAusgewählt = lAusgewählt + 1
        End If
        
        If lAusgewählt > 1 Then
            anzeige "normal", lAusgewählt & " Kunden sind ausgewählt.", Label18
        ElseIf lAusgewählt = 1 Then
            anzeige "normal", lAusgewählt & " Kunde ist ausgewählt.", Label18
        Else
            anzeige "normal", "", Label18
        End If
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub KUTEILMEupdate()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow    As Long
    Dim lRows   As Long
    Dim lcol    As Long
    Dim cKdnr As String
    Dim sSQL As String
    
    
    MSHFLEX1.Redraw = False
    
    lRows = MSHFLEX1.Rows
    lRows = lRows - 1
    lcol = 0
    
    For lrow = 2 To lRows
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        If MSHFLEX1.Text = "" Then
            MSHFLEX1.Col = 1
            cKdnr = MSHFLEX1.Text
            If IsNumeric(cKdnr) Then
                sSQL = "Delete from KUTEILME where knummer = " & cKdnr
                gdBase.Execute sSQL, dbFailOnError
            End If
        End If
    Next lrow
    
    MSHFLEX1.Redraw = True
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KUTEILMEupdate"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(txtUmsatzVon.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = vbWhite
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtBowertBis_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtBowertBis.BackColor = glSelBack1
    txtBowertBis.SelStart = 0
    txtBowertBis.SelLength = Len(txtBowertBis.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtBowertBis_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtErtragVon_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtErtragVon.BackColor = glSelBack1
    txtErtragVon.SelStart = 0
    txtErtragVon.SelLength = Len(txtErtragVon.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtErtragVon_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtErtragBis_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtErtragBis.BackColor = glSelBack1
    txtErtragBis.SelStart = 0
    txtErtragBis.SelLength = Len(txtErtragBis.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtErtragBis_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtBowertBis_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtBowertBis_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtErtragBis_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtErtragBis_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtErtragVon_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtErtragVon_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtBowertBis_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtBowertBis.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtBowertBis_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtBowertVon_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtBowertVon.BackColor = glSelBack1
    txtBowertVon.SelStart = 0
    txtBowertVon.SelLength = Len(txtBowertVon.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtBowertVon_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtBowertVon_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtBowertVon_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub txtBowertVon_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtBowertVon.BackColor = vbWhite
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtBowertVon_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtDat1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    txtDat1(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtDat1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtDat2_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    txtDat2(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtDat2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtErtragBis_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtErtragBis.BackColor = vbWhite
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtErtragBis_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtErtragVon_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtErtragVon.BackColor = vbWhite
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtErtragVon_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtFil_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtFil.BackColor = glSelBack1
    txtFil.SelStart = 0
    txtFil.SelLength = Len(txtFil.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtFil_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtFil_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtFil.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtFil_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub txtKauf_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    txtKauf(Index).BackColor = glSelBack1
    txtKauf(Index).SelStart = 0
    txtKauf(Index).SelLength = Len(txtKauf(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKauf_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtDat1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtDat1(Index).BackColor = glSelBack1
    txtDat1(Index).SelStart = 0
    txtDat1(Index).SelLength = Len(txtDat1(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtDat1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtDat2_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtDat2(Index).BackColor = glSelBack1
    txtDat2(Index).SelStart = 0
    txtDat2(Index).SelLength = Len(txtDat2(Index).Text)
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtDat2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtKauf_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    txtKauf(Index).BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKauf_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtKdNrBis_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtKdNrBis.BackColor = glSelBack1
    txtKdNrBis.SelStart = 0
    txtKdNrBis.SelLength = Len(txtKdNrBis.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKdNrBis_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtKdNrBis_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKdNrBis_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtKdNrBis_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtKdNrBis.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKdNrBis_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtKdNrVon_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtKdNrVon.BackColor = glSelBack1
    txtKdNrVon.SelStart = 0
    txtKdNrVon.SelLength = Len(txtKdNrVon.Text)
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKdNrVon_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtKdNrVon_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKdNrVon_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtFil_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR

    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890" & Chr(8)
    
    cZeichen = Chr$(KeyAscii)
    
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtFil_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtKdNrVon_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtKdNrVon.BackColor = vbWhite
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtKdNrVon_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtMerkmal_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtMerkmal.BackColor = glSelBack1
    txtMerkmal.SelStart = 0
    txtMerkmal.SelLength = Len(txtMerkmal.Text)
    
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtMerkmal_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtMerkmal_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtMerkmal.BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtMerkmal_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtOrt_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtOrt.BackColor = glSelBack1
    txtOrt.SelStart = 0
    txtOrt.SelLength = Len(txtOrt.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtOrt_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtOrt_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = gcUPPER & gcLower & Chr(8)
    cZeichen = Chr$(KeyAscii)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtOrt_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtOrt_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdHinzu1_Click
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtOrt_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtMerkmal_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        cmdHinzu2_Click
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtMerkmal_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtOrt_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtOrt.BackColor = vbWhite
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtOrt_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtPlzVon_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtPlzVon.BackColor = glSelBack1
    txtPlzVon.SelStart = 0
    txtPlzVon.SelLength = Len(txtPlzVon.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtPlzVon_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtPlzVon_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtPlzVon.BackColor = vbWhite
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtPlzVon_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub txtUmsatzBis_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtUmsatzBis.BackColor = glSelBack1
    txtUmsatzBis.SelStart = 0
    txtUmsatzBis.SelLength = Len(txtUmsatzBis.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtUmsatzBis_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtUmsatzBis_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtUmsatzBis.BackColor = vbWhite
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtUmsatzBis_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtUmsatzVon_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    If bVorhanden Then
        bAender = True
    End If
    txtUmsatzVon.BackColor = glSelBack1
    txtUmsatzVon.SelStart = 0
    txtUmsatzVon.SelLength = Len(txtUmsatzVon.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtUmsatzVon_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub txtUmsatzVon_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    txtUmsatzVon.BackColor = vbWhite
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "txtUmsatzVon_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kundenanalyse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
