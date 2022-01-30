VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWKL15 
   BackColor       =   &H00C0C000&
   Caption         =   "Wareneingang aus Einzellieferung"
   ClientHeight    =   8625
   ClientLeft      =   2220
   ClientTop       =   2730
   ClientWidth     =   11910
   Icon            =   "frmWKL15.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Frame Frame2 
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
      Height          =   5175
      Left            =   120
      TabIndex        =   161
      Top             =   8040
      Visible         =   0   'False
      Width           =   7095
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   240
         TabIndex        =   162
         Top             =   1800
         Width           =   6615
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   163
         Top             =   4440
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
         Index           =   4
         Left            =   4680
         TabIndex        =   164
         Top             =   4440
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
         Index           =   5
         Left            =   2445
         TabIndex        =   165
         Top             =   4440
         Width           =   2205
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
      Begin sevCommand3.Command Command7 
         Height          =   375
         Index           =   6
         Left            =   5280
         TabIndex        =   167
         Top             =   240
         Width           =   1175
         _ExtentX        =   2064
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
         Caption         =   "Import"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   375
         Index           =   7
         Left            =   6480
         TabIndex        =   170
         Top             =   240
         Width           =   375
         _ExtentX        =   661
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
         Caption         =   "?"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label19 
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
         Height          =   255
         Left            =   240
         TabIndex        =   171
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   240
         TabIndex        =   169
         Top             =   1080
         Width           =   6615
      End
      Begin VB.Label Label17 
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
         Height          =   255
         Left            =   240
         TabIndex        =   168
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C000&
         Caption         =   "vorhandene elektronische Lieferscheine"
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
         Left            =   240
         TabIndex        =   166
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.PictureBox picprogress 
      Height          =   135
      Left            =   8640
      ScaleHeight     =   75
      ScaleWidth      =   2235
      TabIndex        =   119
      Top             =   200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtStatus 
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
      Left            =   6240
      TabIndex        =   118
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame4"
      Height          =   3255
      Left            =   9600
      TabIndex        =   76
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
      Begin sevCommand3.Command cmdGo 
         Height          =   310
         Left            =   10200
         TabIndex        =   96
         Top             =   240
         Width           =   1050
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
         MaskColor       =   16777215
         MenuBackColor   =   16448250
         MenuBackColorChecked=   7323903
         MenuBackColorHover=   10935807
         MenuBorderColor =   8388608
         MenuCheckMarkColorFrom=   16514300
         MenuCheckMarkColorTo=   15462640
         MenuForeColor   =   -2147483640
         MenuForeColorHover=   -2147483640
         ButtonStyle     =   2
         Caption         =   "Suchen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   8760
         MaxLength       =   13
         TabIndex        =   89
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   8040
         MaxLength       =   3
         TabIndex        =   88
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   7320
         MaxLength       =   3
         TabIndex        =   87
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   86
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   240
         MaxLength       =   13
         TabIndex        =   85
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   2280
         MaxLength       =   35
         TabIndex        =   84
         Top             =   960
         Width           =   3375
      End
      Begin sevCommand3.Command Command16 
         Height          =   495
         Left            =   5400
         TabIndex        =   83
         Top             =   5520
         Visible         =   0   'False
         Width           =   2895
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
         Caption         =   "Artikel anlegen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command15 
         Height          =   495
         Left            =   8400
         TabIndex        =   79
         Top             =   5520
         Width           =   2895
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
         Caption         =   "Zuordnen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   5
         Left            =   6840
         TabIndex        =   78
         Top             =   6600
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
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
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   4
         Left            =   9120
         TabIndex        =   77
         Top             =   6600
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlex3 
         Height          =   3495
         Left            =   240
         TabIndex        =   80
         Top             =   1440
         Visible         =   0   'False
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6165
         _Version        =   393216
         FocusRect       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   2
         X1              =   240
         X2              =   11280
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Artikel-Bezeichnung"
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
         Left            =   2280
         TabIndex        =   95
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "EAN-Code / Artnr"
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
         Left            =   240
         TabIndex        =   94
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferanten-Nr"
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
         Left            =   5640
         TabIndex        =   93
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   13
         Left            =   8040
         TabIndex        =   92
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Lief.Best.Nr"
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
         Left            =   8760
         TabIndex        =   91
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "LPZ"
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
         Left            =   7320
         TabIndex        =   90
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "unbekannten EAN - Code zu einem bekannten Artikel zuordnen"
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
         Left            =   240
         TabIndex        =   82
         Top             =   240
         Width           =   10695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   120
         TabIndex        =   81
         Top             =   6120
         Width           =   11295
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame4"
      Height          =   615
      Left            =   11280
      TabIndex        =   69
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   3
         Left            =   9120
         TabIndex        =   72
         Top             =   6600
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   2
         Left            =   6840
         TabIndex        =   71
         Top             =   6600
         Width           =   2280
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
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
      Begin sevCommand3.Command Command14 
         Height          =   495
         Left            =   8400
         TabIndex        =   70
         Top             =   4680
         Width           =   2895
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlex2 
         Height          =   3975
         Left            =   360
         TabIndex        =   75
         Top             =   600
         Visible         =   0   'False
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7011
         _Version        =   393216
         FocusRect       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
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
         Height          =   495
         Left            =   120
         TabIndex        =   74
         Top             =   6120
         Width           =   11295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Artikel anlegen"
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
         Height          =   615
         Left            =   240
         TabIndex        =   73
         Top             =   120
         Width           =   10695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame4"
      Height          =   1215
      Left            =   8760
      TabIndex        =   56
      Top             =   6600
      Visible         =   0   'False
      Width           =   1695
      Begin sevCommand3.Command Command13 
         Height          =   375
         Left            =   840
         TabIndex        =   68
         Top             =   5880
         Visible         =   0   'False
         Width           =   2895
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
         Caption         =   "Artikel zuordnen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   1
         Left            =   6840
         TabIndex        =   62
         Top             =   6600
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
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
      Begin sevCommand3.Command Command11 
         Height          =   615
         Index           =   0
         Left            =   9120
         TabIndex        =   60
         Top             =   6600
         Width           =   2295
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command10 
         Height          =   405
         Left            =   840
         TabIndex        =   59
         Top             =   6840
         Visible         =   0   'False
         Width           =   2895
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
         Caption         =   "Übernahme Protokoll"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command9 
         Height          =   375
         Left            =   840
         TabIndex        =   58
         Top             =   6360
         Visible         =   0   'False
         Width           =   2895
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
         Caption         =   "Verarbeitung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command8 
         Height          =   375
         Left            =   840
         TabIndex        =   57
         Top             =   5400
         Visible         =   0   'False
         Width           =   2895
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
         Caption         =   "Inhalt anzeigen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFLEX1 
         Height          =   3375
         Left            =   360
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   5953
         _Version        =   393216
         FocusRect       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin sevCommand3.Command Command12 
         Height          =   360
         Left            =   360
         TabIndex        =   146
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
         Picture         =   "frmWKL15.frx":0442
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label label0 
         Height          =   135
         Index           =   1
         Left            =   1800
         TabIndex        =   104
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label label0 
         Height          =   135
         Index           =   0
         Left            =   1200
         TabIndex        =   103
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "4."
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
         Index           =   3
         Left            =   360
         TabIndex        =   97
         Top             =   6960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "3."
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
         Left            =   360
         TabIndex        =   66
         Top             =   6480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "2."
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
         Left            =   360
         TabIndex        =   65
         Top             =   6000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "1."
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
         Left            =   360
         TabIndex        =   64
         Top             =   5520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Daten aus dem MDE - Gerät / DESADV (Lieferavis)"
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
         Height          =   615
         Left            =   840
         TabIndex        =   63
         Top             =   240
         Width           =   10095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
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
         Height          =   1215
         Index           =   0
         Left            =   4080
         TabIndex        =   61
         Top             =   5280
         Width           =   7335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'Kein
      Height          =   1815
      Left            =   6240
      TabIndex        =   43
      Top             =   7200
      Visible         =   0   'False
      Width           =   1920
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   16
         Left            =   9480
         TabIndex        =   32
         Top             =   840
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   "<<<"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   20
         Left            =   8640
         TabIndex        =   31
         Top             =   840
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   "F4"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   19
         Left            =   7800
         TabIndex        =   30
         Top             =   840
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   18
         Left            =   6960
         TabIndex        =   29
         Top             =   840
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   14
         Left            =   4440
         TabIndex        =   28
         Top             =   840
         Width           =   2520
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   "Rückgängig"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   13
         Left            =   1920
         TabIndex        =   27
         Top             =   840
         Width           =   2520
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   12
         Left            =   1080
         TabIndex        =   26
         Top             =   840
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   "-"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   855
         Index           =   11
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   "+"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   17
         Left            =   9480
         TabIndex        =   24
         Top             =   0
         Width           =   2175
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   ">>>"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   10
         Left            =   8640
         TabIndex        =   23
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
         Caption         =   "00"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   9
         Left            =   7800
         TabIndex        =   22
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   8
         Left            =   6960
         TabIndex        =   21
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   7
         Left            =   6120
         TabIndex        =   20
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   6
         Left            =   5280
         TabIndex        =   19
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   5
         Left            =   4440
         TabIndex        =   18
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   4
         Left            =   3600
         TabIndex        =   17
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   3
         Left            =   2760
         TabIndex        =   16
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   2
         Left            =   1920
         TabIndex        =   15
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin sevCommand3.Command Command2 
         Height          =   840
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   0
         Width           =   840
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   18
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
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Height          =   7575
      Left            =   0
      TabIndex        =   33
      Top             =   1080
      Width           =   11775
      Begin VB.CheckBox Check9 
         Caption         =   "schneller Scan-Mod"
         Height          =   255
         Left            =   120
         TabIndex        =   172
         Top             =   0
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "mit Zu-/Abgang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   159
         Top             =   4080
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ohne Zu-/Abgang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   158
         Top             =   4320
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CheckBox Check8 
         Caption         =   "halten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check6 
         Caption         =   "VPE/Faktor Vorschlag"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   144
         ToolTipText     =   "Verpackungseinheit oder Inhaltsangabe einer Umverpackung"
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   14
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   143
         Text            =   "Text1"
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
      End
      Begin sevCommand3.Command Command7 
         Height          =   220
         Index           =   3
         Left            =   4200
         TabIndex        =   142
         Top             =   2500
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command7 
         Height          =   220
         Index           =   2
         Left            =   4200
         TabIndex        =   141
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   8.25
            Charset         =   2
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
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text3 
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
         Index           =   0
         Left            =   10200
         TabIndex        =   9
         Top             =   3800
         Width           =   975
      End
      Begin VB.CheckBox Check65 
         Caption         =   "Etiketten nur bei Preisänderungen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   137
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CheckBox Check5 
         Caption         =   "überschreiben"
         Height          =   255
         Left            =   9120
         TabIndex        =   135
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CheckBox Check4 
         Caption         =   "halten"
         Height          =   255
         Left            =   10800
         TabIndex        =   134
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Index           =   13
         Left            =   7920
         MaxLength       =   25
         MultiLine       =   -1  'True
         TabIndex        =   132
         Text            =   "frmWKL15.frx":0AD4
         Top             =   5280
         Width           =   3735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Spezialetikett"
         Height          =   255
         Left            =   6840
         TabIndex        =   121
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Etikett sofort"
         Height          =   255
         Left            =   5280
         TabIndex        =   120
         Top             =   3720
         Width           =   1455
      End
      Begin sevCommand3.Command Command18 
         Height          =   495
         Left            =   7200
         TabIndex        =   117
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
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
         Caption         =   "c"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   615
         Index           =   22
         Left            =   9480
         TabIndex        =   116
         Top             =   120
         Visible         =   0   'False
         Width           =   1680
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
         Caption         =   "K Best"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         ItemData        =   "frmWKL15.frx":0ADA
         Left            =   240
         List            =   "frmWKL15.frx":0ADC
         TabIndex        =   106
         Top             =   5520
         Width           =   7335
      End
      Begin sevCommand3.Command Command6 
         Height          =   495
         Left            =   5040
         TabIndex        =   55
         Top             =   1680
         Visible         =   0   'False
         Width           =   1320
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
         Caption         =   "in Filialen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Left            =   5640
         TabIndex        =   54
         Top             =   4080
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   9600
         MaxLength       =   9
         TabIndex        =   7
         Top             =   2560
         Width           =   1695
      End
      Begin VB.TextBox Text1 
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
         Index           =   3
         Left            =   9600
         MaxLength       =   9
         TabIndex        =   8
         Top             =   3000
         Width           =   1695
      End
      Begin sevCommand3.Command Command7 
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   12
         Top             =   4080
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Ändern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3240
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2280
         Width           =   1335
      End
      Begin sevCommand3.Command Command3 
         Height          =   495
         Left            =   7200
         TabIndex        =   10
         Top             =   4080
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Schließen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   15
         Left            =   3960
         TabIndex        =   11
         Top             =   4080
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
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
         Caption         =   "Speichern"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   4935
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Left            =   7680
         TabIndex        =   2
         Top             =   120
         Width           =   1695
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
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
         ItemData        =   "frmWKL15.frx":0ADE
         Left            =   240
         List            =   "frmWKL15.frx":0AE0
         TabIndex        =   126
         Top             =   5280
         Width           =   7335
      End
      Begin sevCommand3.Command Command0 
         Height          =   360
         Index           =   0
         Left            =   11280
         TabIndex        =   147
         ToolTipText     =   "Kalender"
         Top             =   3800
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
         Caption         =   "Original-EAN"
         Height          =   255
         Left            =   5640
         TabIndex        =   149
         Top             =   0
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "als Lieferantenbestellnummer suchen"
         Height          =   255
         Left            =   2160
         TabIndex        =   115
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   11040
         TabIndex        =   156
         Top             =   3520
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   10800
         TabIndex        =   155
         Top             =   3520
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   10560
         TabIndex        =   154
         Top             =   3520
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   10320
         TabIndex        =   153
         Top             =   3520
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   10080
         TabIndex        =   152
         Top             =   3520
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   9840
         TabIndex        =   151
         Top             =   3520
         Width           =   255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   9600
         TabIndex        =   150
         Top             =   3520
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Undurchsichtig
         Height          =   255
         Index           =   6
         Left            =   11040
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Undurchsichtig
         Height          =   255
         Index           =   5
         Left            =   10800
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Undurchsichtig
         Height          =   255
         Index           =   4
         Left            =   10560
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Undurchsichtig
         Height          =   255
         Index           =   3
         Left            =   10320
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Undurchsichtig
         Height          =   255
         Index           =   2
         Left            =   10080
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Undurchsichtig
         Height          =   255
         Index           =   1
         Left            =   9840
         Top             =   3480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Undurchsichtig
         Height          =   255
         Index           =   0
         Left            =   9600
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Zentriert
         BackColor       =   &H000040C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   145
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Left            =   9120
         MouseIcon       =   "frmWKL15.frx":0AE2
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   140
         ToolTipText     =   "mit Doppelklick zur Auswertung"
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   10
         Left            =   11400
         TabIndex        =   139
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "EX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   9
         Left            =   600
         TabIndex        =   138
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "AM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   10440
         TabIndex        =   136
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Notizen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   7920
         TabIndex        =   133
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Stück"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   7
         Left            =   4680
         MouseIcon       =   "frmWKL15.frx":0DEC
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   129
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Filiale"
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
         Height          =   375
         Index           =   43
         Left            =   6720
         TabIndex        =   128
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lbl6 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "l. WE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   6960
         TabIndex        =   127
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   9600
         TabIndex        =   125
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "in Bestellung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   11
         Left            =   5640
         TabIndex        =   124
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Height          =   375
         Index           =   6
         Left            =   5640
         TabIndex        =   123
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
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
         Index           =   17
         Left            =   2160
         TabIndex        =   122
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   5280
         TabIndex        =   114
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "0 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   113
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Summe VK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   112
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Summe REK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   111
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "zuletzt gebuchte Wareneingänge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   107
         Top             =   5040
         Width           =   3735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   105
         Top             =   4680
         Width           =   11415
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Schnitt EK:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   8880
         TabIndex        =   102
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Nettospanne:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   8880
         TabIndex        =   101
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000A&
         Height          =   255
         Left            =   10200
         TabIndex        =   100
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   10200
         TabIndex        =   99
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   240
         X2              =   11640
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "neuer K-VK:"
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
         Index           =   12
         Left            =   8040
         TabIndex        =   53
         Top             =   2660
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EK-Preis:"
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
         Index           =   11
         Left            =   8160
         TabIndex        =   52
         Top             =   3100
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "LiefBestNr:"
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
         Index           =   10
         Left            =   240
         TabIndex        =   50
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Min.Bestand:"
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
         Index           =   1
         Left            =   0
         TabIndex        =   49
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   9600
         TabIndex        =   48
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "alter K-Vk:"
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
         Index           =   9
         Left            =   8040
         TabIndex        =   47
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "unbekannt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   4680
         TabIndex        =   46
         Top             =   720
         Width           =   6255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Lieferant:"
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
         Index           =   8
         Left            =   600
         TabIndex        =   45
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Zu-/Abgang:"
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
         Left            =   0
         TabIndex        =   42
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Listen-Vk:"
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
         Index           =   5
         Left            =   7920
         TabIndex        =   41
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0,00 Euro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   9600
         TabIndex        =   40
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Bestand:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "unbekannt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   38
         Top             =   1200
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   37
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "Artikel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   960
         TabIndex        =   36
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   35
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "EAN / ArtNr:"
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
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5160
      TabIndex        =   108
      Top             =   600
      Width           =   855
   End
   Begin sevCommand3.Command Command2 
      Height          =   360
      Index           =   21
      Left            =   9960
      TabIndex        =   148
      Top             =   480
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   635
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
      ToolTip         =   "Hier können Sie sich die Bildschirmtastatur ein- bzw. ausblenden."
      ToolTipTitle    =   "Tastatur"
      ButtonStyle     =   2
      Caption         =   ""
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command4 
      Height          =   375
      Left            =   9480
      TabIndex        =   130
      ToolTipText     =   "Kalender"
      Top             =   480
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
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Datei"
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
      Index           =   18
      Left            =   10440
      MouseIcon       =   "frmWKL15.frx":10F6
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   160
      ToolTipText     =   "mit Doppelklick"
      Top             =   600
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   10920
      Y1              =   960
      Y2              =   960
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
      Left            =   7680
      TabIndex        =   131
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   6120
      TabIndex        =   110
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnungsrabatt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   109
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Lieferschein - Nr.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   98
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   11160
      MouseIcon       =   "frmWKL15.frx":1400
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "frmWKL15.frx":170A
      ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem MDE - Gerät einlesen möchten"
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Wareneingang aus Einzellieferung"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   240
      TabIndex        =   51
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmWKL15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSpaltenname() As String
Dim sSpaltenbez() As String
Dim aBreite() As Integer
Dim byAnzahlSpalten As Byte
Dim lSelect As Long
Dim newArtikel As ArtikelTyp

Dim bfoundauto As Boolean
Dim fromMde As Boolean
Dim bscanner As Boolean
Dim bLieferavis As Boolean
Dim SpaltennummerMENGE As Byte
Dim iFaktor_Umverpack As Integer
Dim sNeuerEK As String
Private Sub PositionierenWKL15()
    On Error GoTo LOKAL_ERROR
    
    Frame1.Top = 6000
    Frame1.Left = 0
    Frame1.Height = 1815
    Frame1.Width = 11760
    
    With Frame2
        .Top = 1920
        .Left = 2520
        .Height = 5175
        .Width = 7095
    End With
    
    Frame3.Top = 1080
    Frame3.Left = 0
    Frame3.Height = 7575
    Frame3.Width = 12000
    
'    Frame4.Top = 1080
    Frame4.Top = 0
    Frame4.Left = 0
    Frame4.Height = 8160
    Frame4.Width = 12000
    
    Frame6.Top = 840
    Frame6.Left = 0
    Frame6.Height = 8160
    Frame6.Width = 12000
    
    Frame7.Top = 0
    Frame7.Left = 0
    Frame7.Height = 8160
    Frame7.Width = 12000

    MSHFLEX1.Height = 3975
    MSHFLEX1.Left = 240
    MSHFLEX1.Top = 600
    MSHFLEX1.Width = 11055
    
    MSHFlex2.Height = 3975
    MSHFlex2.Left = 240
    MSHFlex2.Top = 600
    MSHFlex2.Width = 11055
    
    
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeereDialogWKL15()
    On Error GoTo LOKAL_ERROR
    
    Text1(0).Text = ""
    Text1(1).Text = gsWeEinzMe
    Text1(5).Text = ""
    Text1(6).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text3(0).Text = ""
    
    If Check8.Value = vbUnchecked Then
        Text1(4).Text = ""
        Label2(4).Caption = ""
    End If

    Label2(0).Caption = "unbekannt"
    Label2(1).Caption = "0"
    Label2(2).Caption = "0"
    Label2(3).Caption = "0,00 " & gcWaehrung
    Label2(5).Caption = "0,00 " & gcWaehrung
    Label12.Caption = ""
    Label13.Caption = ""
    Label3.Caption = "0"
    Label1(17).Caption = ""
    Label2(2).ForeColor = glS1
    Label2(2).BackColor = glH1
    
    Label4(39).Caption = ""
    Label4(39).Visible = False
    
    If Check4.Value = vbUnchecked Then
        Text1(13).Text = ""
    End If
    
    Label2(8).Caption = ""
    Label2(9).Caption = ""
    Label2(10).Caption = ""
    
    sNeuerEK = ""
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeereDialogWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LeseLieferantenPreisWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cLinr As String
    Dim cArtNr As String
    Dim dLEKPR As Double
    
    cArtNr = Label2(2).Caption
    cLinr = Trim$(Str$(Val(Text1(4).Text)))
    
    cSQL = "Select * from ARTLIEF where LINR = " & cLinr & " and ARTNR = " & cArtNr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!lekpr) Then
            dLEKPR = rsrs!lekpr
        Else
            dLEKPR = 0
        End If
    Else
        dLEKPR = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If dLEKPR <> 0 Then
        Text1(3).Text = Format$(dLEKPR, "#####0.00")
    Else
        Text1(3).Text = ""
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LeseLieferantenPreisWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SchreibeDatenWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim lMinBest       As Long
    Dim lLpz           As Long
    Dim lHeute         As Long
    Dim ctmp           As String
    Dim cSQL           As String
    Dim cMeld          As String
    Dim cArtNr         As String
    Dim cBezeich       As String
    Dim cEtiMerk       As String
    Dim cLiBesNr       As String
    Dim cLiBesNrDialog As String
    Dim cEkPr          As String
    Dim cLinr          As String
    Dim cLiNrDialog    As String
    Dim cJetzt         As String
    Dim cEAN           As String
    
    Dim cPreis         As String
    Dim sTab           As String
    Dim cLS            As String
    
    Dim dVkPr          As Double
    Dim dKVkPr1        As Double
    Dim dBestand       As Double
    Dim dPreis         As Double
    Dim dEkpr          As Double
    Dim dEkPrAlt       As Double
    Dim dEkPrSchnitt   As Double
    Dim dWertAlt       As Double
    Dim dWertNeu       As Double
    Dim dWert          As Double
    Dim dGewichtWert   As Double
    Dim dStückWert     As Double
    Dim dAlt           As Double
    Dim iArtAnzahl     As Integer
    Dim iFehlerstufe   As Integer
    Dim iRet           As Integer
    Dim bNeu           As Boolean
    Dim bTrans         As Boolean
    
    Dim rsA             As Recordset
    Dim rsZ             As Recordset
    Dim rsrs            As Recordset
    Dim rsEti           As Recordset
    Dim rsHis           As Recordset
    Dim rsHis1          As Recordset
    Dim rsArtlief       As Recordset
    Dim rsZutemp        As Recordset
    
    Dim i               As Integer
    Dim iZBestand       As Integer
    Dim siRechrabatt    As Single
    Dim lMDHDAT         As Long
    Dim lKJADate        As Long
    Dim cKJAZeit        As String
    
    lKJADate = Fix(Now)
    cKJAZeit = Format$(Now, "HH:MM:SS")

    cArtNr = Label2(2).Caption
    
    ctmp = Trim$(Text1(1).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    
    If Label2(7).Caption = "Stück" Then
        dStückWert = Val(ctmp)
    ElseIf Label2(7).Caption = "Kg" Then
        dGewichtWert = CDbl(ctmp)
    End If
    
    ctmp = Trim$(Text1(2).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    dWert = Val(ctmp)
    
    iFehlerstufe = 0
    bTrans = False
    
    If Combo1.Text <> "" And Not Combo1.Text = "keine" Then
        cLS = Trim(Combo1.Text)
    Else
        cLS = ""
    End If
   
    If Text1(4).Text <> "" Then
        cLiNrDialog = Text1(4).Text
    Else
        MsgBox "Die Lieferantennummer fehlt", vbInformation, "Winkiss Hinweis:"
        Text1(4).SetFocus
        Exit Sub
    End If

    cLiBesNrDialog = Text1(6).Text
    
    Dim cPreis1 As String
    
    cPreis = Text1(2).Text
    cPreis = Trim$(cPreis)
    If cPreis <> "" Then
        If InStr(cPreis, ",") = 0 Then
            cPreis = Format$((Val(cPreis) / 100), "#####0.00")
        End If
        cPreis1 = cPreis
        cPreis = fnMoveComma2Point$(cPreis)
        
        dPreis = Val(cPreis)
        If dPreis > 100000 Then
            MsgBox "Der eingegebene Preis ist zu hoch!", vbInformation, "Winkiss Hinweis:"
            Text1(2).SetFocus
            Exit Sub
        End If
    End If
    
    cEkPr = Text1(3).Text
    cEkPr = Trim$(cEkPr)
    If cEkPr <> "" Then
        If InStr(cEkPr, ",") = 0 Then
            cEkPr = Format$((Val(cEkPr) / 100), "#####0.00")
        End If
        cEkPr = fnMoveComma2Point$(cEkPr)
        dEkpr = Val(cEkPr)
        If dEkpr > 100000 Then
            MsgBox "Der eingegebene Preis ist zu hoch!", vbInformation, "Winkiss Hinweis:"
            Text1(3).SetFocus
            Exit Sub
        End If
    Else
        dEkpr = 0
    End If
    lMinBest = Val(Text1(5).Text)

    
    If Abs(dStückWert) > 99 Then
        If bLieferavis = True Then
            iRet = vbYes
        Else
            iRet = MsgBox("Mengenangabe von " & dStückWert & " fraglich! Trotzdem speichern?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        End If
        
        If iRet = vbNo Then
            Text1(1).SetFocus
            Exit Sub
        End If
    End If
    
    If Abs(dStückWert) > 9999 Then
        iRet = MsgBox("Mengen über 9999 werden nicht gespeichert!", vbCritical, "Winkiss Abbruch:")
        Text1(1).SetFocus
        Exit Sub
    End If
    
    cArtNr = Label2(2).Caption
    cArtNr = Trim$(cArtNr)
    If cArtNr = "" Then
        MsgBox "Artikel-Nr fehlt! Daten speichern nicht möglich!", vbCritical, "Winkiss Abbruch:"
        Text1(0).SetFocus
        Exit Sub
    End If
    
    iFehlerstufe = 2
    cSQL = "Select * from ARTIKEL where ARTNR = " & cArtNr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)

    If Not rsrs.EOF Then
        rsrs.MoveLast
        If rsrs.RecordCount > 1 Then
            MsgBox "Mehr als 1 Artikeleintrag gefunden! Daten speichern nicht möglich!", vbCritical, "Winkiss Abbruch:"
            Text1(0).SetFocus
            Exit Sub
        End If
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!BEZEICH) Then
            cBezeich = rsrs!BEZEICH
        Else
            cBezeich = ""
        End If
        If Not IsNull(rsrs!vkpr) Then
            dVkPr = rsrs!vkpr
        Else
            dVkPr = 0
        End If
        
        If Not IsNull(rsrs!KVKPR1) Then
            dKVkPr1 = rsrs!KVKPR1
        Else
            dKVkPr1 = 0
        End If
        If Not IsNull(rsrs!BESTAND) Then
            dAlt = rsrs!BESTAND
        Else
            dAlt = 0
        End If
        If Not IsNull(rsrs!ETIMERK) Then
            cEtiMerk = rsrs!ETIMERK
        Else
            cEtiMerk = ""
        End If
        If Not IsNull(rsrs!LIBESNR) Then
            cLiBesNr = rsrs!LIBESNR
        Else
            cLiBesNr = ""
        End If
        If Not IsNull(rsrs!EAN) Then
            cEAN = rsrs!EAN
        Else
            cEAN = ""
        End If
        
        If Not IsNull(rsrs!linr) Then
            cLinr = rsrs!linr
        Else
            cLinr = ""
        End If
        
        If Not IsNull(rsrs!LPZ) Then
            lLpz = rsrs!LPZ

        Else
            lLpz = 0
        End If
        
        iFehlerstufe = 33
        
        Dim cekalt1 As String
        If Not IsNull(rsrs!ekpr) Then
            cekalt1 = rsrs!ekpr
            If IsNumeric(cekalt1) Then
                dEkPrAlt = CDbl(cekalt1)
            Else
                dEkPrAlt = 0
            End If
            
        Else
            dEkPrAlt = 0
        End If
        
        
        iFehlerstufe = 34
        
        If Val(Combo2.Text) > 0 Then
            siRechrabatt = Combo2.Text
        End If
        
        iFehlerstufe = 35
        
        If dEkPrAlt <> 0 And dEkpr <> 0 Then
            If dEkPrAlt / dEkpr > 10 Or dEkpr / dEkPrAlt > 10 Then
                cMeld = "Hoher Unterschied im EK-Preis festgestellt!" & vbCrLf
                cMeld = cMeld & "bisheriger Schnitt = " & Format$(dEkPrAlt, "###,##0.00") & vbCrLf
                cMeld = cMeld & "neuer EK-Preis = " & Format$(dEkpr, "###,##0.00") & vbCrLf
                cMeld = cMeld & "Trotzdem speichern?"
                
                
                If bLieferavis = True Then
                    iRet = vbYes
                Else
                    iRet = MsgBox(cMeld, vbYesNo + vbQuestion, "ABWEICHUNG")
                End If
                
                If iRet <> vbYes Then
                    Exit Sub
                End If
            End If
            
            iFehlerstufe = 43
            
            'Schnitt beginn
            Dim zugang As Long
            Dim bestandalt As Long
            
            If Val(Combo2.Text) > 0 Then
                siRechrabatt = Combo2.Text
            Else
                siRechrabatt = 0
            End If
            
            If Val(siRechrabatt) > 0 Then
                dEkpr = dEkpr - ((dEkpr * siRechrabatt) / 100)
            End If
            
            iFehlerstufe = 43

            zugang = CLng(dStückWert)
            bestandalt = CLng(Label2(1).Caption)
            iFehlerstufe = 44
            dWertNeu = SchnittEKBerechnung(cArtNr, CLng(Text1(4).Text), zugang, dEkpr, bestandalt)
            'Schnitt ende
        Else
            If dEkpr >= 100000 Then
                MsgBox "Der eingegebene Preis ist zu hoch!", vbCritical, "STOP!"
                Text1(3).SetFocus
                Exit Sub
            End If
            
            'Schnitt beginn
            If Val(Combo2.Text) > 0 Then
                siRechrabatt = Combo2.Text
            Else
                siRechrabatt = 0
            End If
            
            If Val(siRechrabatt) > 0 Then
                dEkpr = dEkpr - ((dEkpr * siRechrabatt) / 100)
                
            End If
            
            If IsNumeric(Text1(1).Text) Then
                zugang = CLng(Text1(1).Text)
            Else
                zugang = 0
            End If
            
            bestandalt = CLng(Label2(1).Caption)
            
            iFehlerstufe = 55
            
            dWertNeu = SchnittEKBerechnung(cArtNr, CLng(Text1(4).Text), zugang, dEkpr, bestandalt)
            'Schnitt ende
        End If
        
        iFehlerstufe = 36
        
        lHeute = Fix(Now)
        cJetzt = Format$(Now, "HH:MM")
        
        If Not tableSuchenDBKombi("ZuTemp", 2) Then
            cSQL = "Create Table ZuTemp "
            cSQL = cSQL & "( "
            cSQL = cSQL & "ARTNR LONG"
            cSQL = cSQL & ", BEZEICH Text (35) "
            cSQL = cSQL & ", EAN TEXT (13) "
            cSQL = cSQL & ", LINR LONG "
            cSQL = cSQL & ", ADATE DATETIME "
            cSQL = cSQL & ", UHRZEIT TEXT (5) "
            cSQL = cSQL & ", BEDNU long "
            cSQL = cSQL & ", BEDNAME TEXT (32) "
            cSQL = cSQL & ", FILIALNR BYTE "
            cSQL = cSQL & ", BESTANDALT INTEGER "
            cSQL = cSQL & ", BEWEGUNG INTEGER "
            cSQL = cSQL & ", BESTANDNEU INTEGER "
            cSQL = cSQL & ", EKPR SINGLE "
            cSQL = cSQL & ", VKPR SINGLE "
            cSQL = cSQL & ", KVKPR1 SINGLE "
            cSQL = cSQL & ", LIBESNR TEXT (13) "
            cSQL = cSQL & ", LS TEXT (20) "
            cSQL = cSQL & ") "
            gdApp.Execute cSQL, dbFailOnError
        Else
            If Not SpalteInTabellegefundenNEW("ZuTemp", "VKPR", gdApp) Then
                SpalteAnfuegenNEW "ZuTemp", "VKPR", "SINGLE", gdApp
            End If
            
            If Not SpalteInTabellegefundenNEW("ZuTemp", "KVKPR1", gdApp) Then
                SpalteAnfuegenNEW "ZuTemp", "KVKPR1", "SINGLE", gdApp
            End If
            
            If Not SpalteInTabellegefundenNEW("ZuTemp", "LS", gdApp) Then
                SpalteAnfuegenNEW "ZuTemp", "LS", "TEXT (20)", gdApp
            End If
            
         
        End If
               
        iFehlerstufe = 37
               
        Set rsZutemp = gdApp.OpenRecordset("Zutemp", dbOpenTable)
        
        cSQL = "Select * from ZUGANG where ARTNR = -1"
        Set rsHis = gdBase.OpenRecordset(cSQL)
        
        cSQL = "Select * from ZUGANGF where ARTNR = -1"
        Set rsHis1 = gdBase.OpenRecordset(cSQL)
        
        cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLiNrDialog & " "
        Set rsArtlief = gdBase.OpenRecordset(cSQL)
        
        '***** Anfang TRANSAKTIONS-Klammerung *****
        If islocked(rsrs) Then
            err.Number = 3008
            GoTo LOKAL_ERROR
        End If
        
        iFehlerstufe = 38
        
        BeginTrans
        bTrans = True
        
        If dStückWert <> 0 Then
            Bestandsveraenderung cArtNr, CLng(dAlt + dStückWert), "WE / Einzellieferung"
            ABinBESTAKT cArtNr, CLng(dStückWert), "WE / Einzellieferung"
        End If
        
        rsrs.Edit
        rsrs!SYNStatus = "E"
        
        iFehlerstufe = 381
        
        Dim bKVKPR1     As Boolean
        bKVKPR1 = False

        If cPreis1 <> "" Then
            'Hat sich der KVKPR1 geändert
    
            If Not IsNull(rsrs!KVKPR1) Then
                If Trim(CStr(rsrs!KVKPR1)) <> Trim$(cPreis1) Then
                    bKVKPR1 = True
                End If
            Else
                bKVKPR1 = True
            End If
        
            Dim dlinr1 As Long
            If Not IsNull(rsrs!linr) Then
                dlinr1 = rsrs!linr
            Else
                dlinr1 = 0
            End If
            
            iFehlerstufe = 382
            
            If Val(Text1(4).Text) = Val(dlinr1) Then
                Dim se      As String
                Dim sMW     As String
                
                If Not IsNull(rsrs!MWST) Then
                    sMW = rsrs!MWST
                Else
                    sMW = "V"
                End If
                
                iFehlerstufe = 383
                
                If rsArtlief.RecordCount = 0 Then
                   se = "0"
                Else
                    If gsSpanne = "LEK" Then
                        If Not IsNull(rsArtlief!lekpr) Then
                            se = rsArtlief!lekpr
                        Else
                            se = "0"
                        End If
                        
                    iFehlerstufe = 384
                    
                    ElseIf gsSpanne = "SEK" Then
                        If Not IsNull(rsrs!ekpr) Then
                            se = rsrs!ekpr
                        Else
                            se = "0"
                        End If
                        
                    End If
                End If
                
                rsrs!SPANNE = NettospanneInProzent(Trim(Text1(2).Text), se, sMW)
            End If
        End If
        
        If cEkPr <> "" Then
            rsrs!ekpr = dWertNeu
        End If
        
        rsrs!MINBEST = lMinBest
        
        If rsrs!GEFUEHRT = "N" Then
            
            If gbWEautoGef Then
                rsrs!GEFUEHRT = "J"
                insertArtikelDetail lKJADate, cKJAZeit, gcKasNum, CInt(gcBedienerNr), CLng(cArtNr), "gefuehrt", "J"
            End If
        End If
        
        If Text3(0).Text <> "" Then
            lMDHDAT = DateValue(Text3(0).Text)
            insertArtikelMDH lKJADate, cKJAZeit, CInt(gcBedienerNr), CLng(cArtNr), lMDHDAT
        End If
        
        
        rsrs!GEFUEHRT = "J"
        
        If cLinr = cLiNrDialog Then
            rsrs!LIBESNR = cLiBesNrDialog
        End If
        If gsArtikelFarbe <> "" Then
            rsrs!AWM = Label4(32).Tag
        End If
        
        iFehlerstufe = 4
        
        rsrs!NOTIZEN = Trim(Text1(13).Text)
        rsrs!LASTDATE = DateValue(Now)
        rsrs!LASTTIME = TimeValue(Now)
        rsrs.Update
        
        iFehlerstufe = 5
        rsHis.AddNew
        rsHis!artnr = Val(cArtNr)
        iFehlerstufe = 501
        rsHis!BEZEICH = cBezeich
        iFehlerstufe = 502
        rsHis!linr = Val(cLiNrDialog)
        iFehlerstufe = 503
        rsHis!EAN = cEAN
        iFehlerstufe = 504
        rsHis!ADATE = lHeute
        iFehlerstufe = 505
        rsHis!Uhrzeit = cJetzt
        iFehlerstufe = 506
        rsHis!BEDNU = Val(gcBedienerNr)
        iFehlerstufe = 507
        rsHis!bedname = gcUserName
        iFehlerstufe = 508
        rsHis!FILIALNR = 1
        iFehlerstufe = 509
        rsHis!bestandalt = dAlt
        iFehlerstufe = 510
        rsHis!BEWEGUNG = dStückWert
        iFehlerstufe = 511
        rsHis!BESTANDneu = dAlt + dStückWert
        iFehlerstufe = 512
        rsHis!ekpr = Val(cEkPr)
        iFehlerstufe = 513
        rsHis!LS = cLS
        iFehlerstufe = 514
        rsHis!rek = dEkpr
        rsHis.Update
        
        
        rsHis1.AddNew
        rsHis1!artnr = Val(cArtNr)
        rsHis1!BEZEICH = cBezeich
        rsHis1!linr = Val(cLiNrDialog)
        rsHis1!EAN = cEAN
        rsHis1!ADATE = lHeute
        rsHis1!Uhrzeit = cJetzt
        rsHis1!BEDNU = Val(gcBedienerNr)
        rsHis1!bedname = gcUserName
        rsHis1!FILIALNR = gcFilNr
        rsHis1!bestandalt = dAlt
        rsHis1!BEWEGUNG = dStückWert
        rsHis1!BESTANDneu = dAlt + dStückWert
        rsHis1!ekpr = Val(cEkPr)
        rsHis1!LS = cLS
        rsHis1!rek = dEkpr
        rsHis1!SENDOK = False
        rsHis1.Update
        
        rsZutemp.AddNew
        rsZutemp!artnr = Val(cArtNr)
        rsZutemp!BEZEICH = cBezeich
        rsZutemp!linr = Val(cLiNrDialog)
        rsZutemp!EAN = cEAN
        rsZutemp!ADATE = lHeute
        rsZutemp!Uhrzeit = cJetzt
        rsZutemp!BEDNU = Val(gcBedienerNr)
        rsZutemp!bedname = gcUserName
        rsZutemp!FILIALNR = 1
        rsZutemp!bestandalt = dAlt
        rsZutemp!BEWEGUNG = dStückWert
        rsZutemp!BESTANDneu = dAlt + dStückWert
        rsZutemp!ekpr = Val(cEkPr)
        rsZutemp!LIBESNR = cLiBesNrDialog
        
        rsZutemp!vkpr = dVkPr
        rsZutemp!KVKPR1 = dKVkPr1
        rsZutemp!LS = cLS
        rsZutemp.Update
    
        If KundenbestBestätigung(cArtNr, dStückWert) = True Then
            Command2(22).Visible = True
        End If
        
        newArtikel.artnr = Val(cArtNr)
        newArtikel.BEZEICH = cBezeich
        newArtikel.ZubuchMe = dStückWert
        newArtikel.REKPR = dEkpr
        newArtikel.lekpr = Val(cEkPr)
        newArtikel.ekpr = dWertNeu
        newArtikel.KVKPR1 = cPreis1 'Text1(2).Text
        
        iFehlerstufe = 6
        If rsArtlief.EOF Then
            rsArtlief.AddNew
            rsArtlief!SYNStatus = "A"
        Else
            rsArtlief.Edit
            rsArtlief!SYNStatus = "E"
        End If
        
        rsArtlief!artnr = Val(cArtNr)
        rsArtlief!linr = Val(cLiNrDialog)
        rsArtlief!lekpr = Val(cEkPr)
        rsArtlief!LIBESNR = cLiBesNrDialog
        rsArtlief.Update
        
        CommitTrans
        
        rsHis.Close: Set rsHis = Nothing
        rsHis1.Close: Set rsHis1 = Nothing
        rsZutemp.Close: Set rsZutemp = Nothing
        
        bTrans = False
        '***** Ende TRANSAKTIONS-Klammerung *****
        rsrs.Close: Set rsrs = Nothing
        
        If bKVKPR1 Then
            Artikelveraenderung cArtNr, Trim$(cPreis1), "WE / Einzellieferung", "KVKPR1"
        End If
        
        iFehlerstufe = 7
        
        If Frame4.Visible = True Then
        
        Else
            anzeige "LASER", "", Label14
        End If
    Else
        MsgBox "Keine Artikeldaten gefunden! Daten speichern nicht möglich!", vbCritical, "Winkiss Hinweis:"
        Text1(0).SetFocus
        Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    
    
    If Trim$(Text1(2).Text) <> Trim$(Left(Label2(5).Caption, Len(Label2(5).Caption) - 4)) Then
'        If (Trim$(Label2(1).Caption) <> dStückWert) And (dStückWert <> 0) Then
        If dStückWert > 0 Then
            '** es gibt neue KVKPR1 UND neuer BESTAND **
            dBestand = dAlt + dStückWert '**  ARTIKEL **
            If gcFilNr = "1" Then
                
                For i = 1 To giAnzFil
                    If i <> 1 Then
                        If gbETIONLYME = False Then
                            iZBestand = ermBestandfromZbestand(cArtNr, i)
                            schreibeWKEtidru cArtNr, CLng(iZBestand), CLng(i)
                        End If
                    Else
                        schreibeWKEtidru cArtNr, CLng(dBestand), CLng(i)
                    End If
                    
                    If Combo1.Text <> "" And Not Combo1.Text = "keine" Then
                        sTab = "ETIDRULS" 'etidruls füllen
                    End If
                Next i
            Else
                If Combo1.Text <> "" And Not Combo1.Text = "keine" Then
                    sTab = "ETIDRULS"            'etidruls füllen
                Else
                    schreibeWKEtidru cArtNr, CLng(dBestand), CLng(gcFilNr)
                End If
            End If
        Else '** nur Preisveränderung,kein Bestandveränderung **
            If gcFilNr = "1" Then
                dBestand = dAlt + dStückWert '**  ARTIKEL **
                For i = 1 To giAnzFil
                    If i <> 1 Then
                        If gbETIONLYME = False Then
                            iZBestand = ermBestandfromZbestand(cArtNr, i)
                            schreibeWKEtidru cArtNr, CLng(iZBestand), CLng(i)
                        End If
                    Else
                        schreibeWKEtidru cArtNr, CLng(dBestand), CLng(i)
                    End If
                    
                    If Combo1.Text <> "" And Not Combo1.Text = "keine" Then
                        sTab = "ETIDRULS"            'etidruls füllen
                    End If
                Next i
            Else
                If Combo1.Text <> "" And Not Combo1.Text = "keine" Then
                    sTab = "ETIDRULS"            'etidruls füllen
                Else
                
                    If Option1(0).Value = True Then
                        schreibeWKEtidru cArtNr, CLng(dBestand), CLng(gcFilNr)
                    
                    ElseIf Option1(1).Value = True Then
                        schreibeWKEtidru cArtNr, 1, CLng(gcFilNr)
                    End If
                    
                    
                End If
            End If
        
        End If
    Else '** nur Bestandveränderung,kein Preisveränderung  **
        If Combo1.Text <> "" And Not Combo1.Text = "keine" Then
            sTab = "ETIDRULS"            'etidruls füllen
        Else
            If Check65.Value = vbUnchecked Then 'Einstellung: Etiketten nur bei KVK Änderung schreiben
            
                schreibeWKEtidru cArtNr, CLng(dStückWert), CLng(gcFilNr)

            End If
        End If
    End If
    
    If Combo1.Text <> "" And Not Combo1.Text = "keine" Then
        cSQL = "Insert Into ETIDRULS "
        cSQL = cSQL & "( ARTNR, BEZEICH, BESTAND "
        cSQL = cSQL & ", ANZAHL, VKPR, LIBESNR "
        cSQL = cSQL & ", EAN, LPZ, LINR "
        cSQL = cSQL & ", FILNR, WEDate, LS "
        cSQL = cSQL & ") "
        cSQL = cSQL & " values "
        cSQL = cSQL & "( "
        cSQL = cSQL & " '" & cArtNr & "' "
        cSQL = cSQL & ", '" & cBezeich & "' "
        cSQL = cSQL & ", '" & dBestand & "' "
        cSQL = cSQL & ", '" & dStückWert & "' "
        cSQL = cSQL & ", '" & Val(cPreis) & "' "
        cSQL = cSQL & ", '" & cLiBesNr & "' "
        cSQL = cSQL & ", '" & cEAN & "' "
        cSQL = cSQL & ", '" & lLpz & "' "
        cSQL = cSQL & ", '" & Val(cLinr) & "' "
        cSQL = cSQL & ", '" & gcFilNr & "' "
        cSQL = cSQL & ", '" & DateValue(Now) & "' "
        cSQL = cSQL & ", '" & Trim(Combo1.Text) & "' "
        cSQL = cSQL & ") "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    If Check2.Value = vbChecked Or Check3.Value = vbChecked Then
        setzedrucker gcEtikettenDrucker
    End If
    
    If Check2.Value = vbChecked Then
        'Drucketiketten
        If Modul6.FindFile(gcDBPfad, "aWKL30ys.rpt") Then
            SchreibeEtiforSofort "aWKL30ys"
        ElseIf Modul6.FindFile(App.Path, "aWKL30xs.rpt") Then
            SchreibeEtiforSofort "aWKL30xs"
        
        End If
        
    End If
    
    If Check3.Value = vbChecked Then
    
        If Modul6.FindFile(App.Path, "aWKL30zs.rpt") Then
            SchreibeEtiforSofort "aWKL30zs"
        Else
            'Drucketiketten
            For i = 1 To Val(dStückWert)
                DruckeLFNREtikett
            Next i
        End If
    End If
    
    If Check2.Value = vbChecked Or Check3.Value = vbChecked Then
        setzedrucker gcListenDrucker
    End If
    
    LeereDialogWKL15
    SchreibeListe
    Text1(0).SetFocus

Exit Sub
LOKAL_ERROR:
    
    
    If err.Number = 3008 Or err.Number = 3218 Then 'Datenbanksperrung
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
                Command2_Click 15
            Else
                Exit Sub
            End If
        
        Else
            Command2_Click 15 'nochmal
        End If
    Else
        If bTrans Then
            Rollback
        End If
            
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeDatenWKL15"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten." & Trim$(Str$(iFehlerstufe))
        
        Fehlermeldung1
        
        Resume Next 'bleibt 05.04.04
    End If
End Sub
Private Sub SchreibeDatenWKL15fürKilo()
    On Error GoTo LOKAL_ERROR
    
    Dim cArtNr          As String
    Dim ctmp            As String
    Dim dGewichtWert    As Double
    Dim cSQL            As String
    Dim rsrs            As Recordset

    cArtNr = Label2(2).Caption
    
    ctmp = Trim$(Text1(1).Text)
    ctmp = fnMoveComma2Point$(ctmp)
    
    dGewichtWert = Val(ctmp)
    
    cSQL = "Select * from KILOART where ARTNR = " & cArtNr & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)

    If Not rsrs.EOF Then
        rsrs.Edit
        If Not IsNull(rsrs!GewichtKG) Then
             rsrs!GewichtKG = rsrs!GewichtKG + dGewichtWert
        End If
        rsrs.Update
    Else
        rsrs.AddNew
        
        rsrs!artnr = cArtNr
        rsrs!GewichtKG = dGewichtWert
        
        rsrs.Update
    End If
    rsrs.Close
    
    newArtikel.artnr = Val(cArtNr)
    newArtikel.BEZEICH = Label2(0).Caption
    newArtikel.ZubuchKg = dGewichtWert
'    newArtikel.REKPR = dEkpr
'    newArtikel.lekpr = Val(cEkPr)
'    newArtikel.ekpr = dWertNeu
'    newArtikel.KVKPR1 = Text1(2).Text
    
    LeereDialogWKL15
    SchreibeListeKg
    Text1(0).SetFocus

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWKL15fürKilo"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SchreibeEtiforSofort(srepname As String)
    On Error GoTo LOKAL_ERROR
        
    Dim lartnr      As Long
    Dim cSQL        As String
    Dim cLinr       As String
    Dim cArtNr      As String
    Dim cAnzahl     As String
    Dim bAnd        As Boolean
    ReDim acArtNr(0 To 0) As String
    ReDim acAnzEti(0 To 0) As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    bAnd = False
    
    cLinr = Text1(4).Text
    cLinr = Trim$(cLinr)
    
    cArtNr = Label2(2).Caption
    cArtNr = Trim$(cArtNr)
    If Not IsNumeric(cArtNr) Then
        If Not IsNumeric(cLinr) Then
            Exit Sub
        Else
        
        End If
    End If
    
    cAnzahl = Text1(1).Text
    cAnzahl = Trim$(cAnzahl)
    If cAnzahl = "" Then cAnzahl = "1"
    
    cSQL = "Delete from ETI" & srechnertab
    gdBase.Execute cSQL, dbFailOnError

    cSQL = " Select "
    cSQL = cSQL & "  ARTIKEL.ARTNR "
    cSQL = cSQL & ", ARTIKEL.BEZEICH "
    cSQL = cSQL & ", ARTIKEL.KVKPR1 as VKPR "
    cSQL = cSQL & ", ARTIKEL.BESTAND "
    cSQL = cSQL & ", ARTIKEL.SYNSTATUS "
    cSQL = cSQL & ", " & cAnzahl & " as Anzahl "
    cSQL = cSQL & ", ARTIKEL.LIBESNR "
    cSQL = cSQL & ", ARTIKEL.EAN "
    cSQL = cSQL & ", ARTIKEL.LPZ "
    cSQL = cSQL & ", ARTIKEL.LINR "
    cSQL = cSQL & ", " & gcFilNr & " as Filnr "
    cSQL = cSQL & " from ARTIKEL inner join ARTLIEF on "
    cSQL = cSQL & " ARTIKEL.ARTNR = ARTLIEF.ARTNR and Artikel.LINR = Artlief.LINR where "
    
    If cLinr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If

        cSQL = cSQL & "ARTLIEF.LINR = " & cLinr & " "
        bAnd = True
    End If
    
    If cArtNr <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        Select Case Len(cArtNr)
            Case Is > 8
                cSQL = cSQL & "ARTIKEL.EAN = '" & cArtNr & "' "
                cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cArtNr & "' "
                cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cArtNr & "' "
            Case Is = 8
                If Left(cArtNr, 1) = "2" Or Left(cArtNr, 1) = "0" Then
                    cArtNr = Mid(cArtNr, 2, 6)
                    cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
                Else
                    cSQL = cSQL & "ARTIKEL.EAN = '" & cArtNr & "' "
                End If
            Case Is = 6
                cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
            Case Else
                cSQL = cSQL & "ARTIKEL.ARTNR = " & cArtNr & " "
        End Select
    End If
    
    Dim rsEti As Recordset
    Dim siAnzeige As Single
    
    siAnzeige = 0
    
    Set rsEti = gdBase.OpenRecordset("ETI" & srechnertab)
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
'            siAnzeige = siAnzeige + 1
'            Label2(7).Caption = siAnzeige
'            Label2(7).Refresh

            rsEti.AddNew
            
            If Not IsNull(rsrs!artnr) Then
                rsEti!artnr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                rsEti!BEZEICH = rsrs!BEZEICH
            End If
            
            If Not IsNull(rsrs!vkpr) Then
                rsEti!vkpr = rsrs!vkpr
            End If
            
            If Not IsNull(rsrs!BESTAND) Then
                rsEti!BESTAND = rsrs!BESTAND
            End If
            
            If Not IsNull(rsrs!SYNStatus) Then
                rsEti!SYNStatus = rsrs!SYNStatus
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                rsEti!ANZAHL = rsrs!ANZAHL
            End If
            
            If Not IsNull(rsrs!LIBESNR) Then
                rsEti!LIBESNR = rsrs!LIBESNR
            End If
            
            If Not IsNull(rsrs!EAN) Then
                rsEti!EAN = rsrs!EAN
            End If
            
            If Not IsNull(rsrs!LPZ) Then
                rsEti!LPZ = rsrs!LPZ
            End If
            
            If Not IsNull(rsrs!linr) Then
                rsEti!linr = rsrs!linr
            End If
            
            If Not IsNull(rsrs!filnr) Then
                rsEti!filnr = rsrs!filnr
            End If
            
            rsEti.Update
        rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close: Set rsrs = Nothing
    rsEti.Close: Set rsEti = Nothing
    

    '************************** ab hier gemeinsame Verarbeitung ****************
    
    cSQL = "Delete from ETI" & srechnertab & " where SYNSTATUS = 'D' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from ETI" & srechnertab & " where anzahl < 1"
    gdBase.Execute cSQL, dbFailOnError
    
    Dim lAnzahl As Long
    Dim lFil As Long
    
    lAnzahl = -1
    
    lFil = CLng(gcFilNr)
    
    
    Set rsrs = gdBase.OpenRecordset("ETI" & srechnertab, dbOpenTable)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!artnr) Then
            cArtNr = rsrs!artnr
            cAnzahl = rsrs!ANZAHL
            lAnzahl = lAnzahl + 1
            ReDim Preserve acArtNr(0 To lAnzahl) As String
            ReDim Preserve acAnzEti(0 To lAnzahl) As String
            acArtNr(lAnzahl) = cArtNr
            acAnzEti(lAnzahl) = cAnzahl
            
        End If
        
        rsrs.MoveNext
        Loop
    Else
        
        MsgBox "Keine Artikel bzw. Artikelbestände zum Speichern gefunden!", vbInformation, "INFO"
        Exit Sub
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If UCase(srepname) = "AWKL30XS" Then
        DruckeGrundPreisEtikettenWKL30kleinspezial acArtNr(), lAnzahl, srepname
    ElseIf UCase(srepname) = "AWKL30YS" Then
        DruckeStrichcodeY acArtNr(), lAnzahl, acAnzEti()
        reportbildschirmToPrinterETI "aWKL30ys", gcEtikettenDrucker, False
    ElseIf UCase(srepname) = "AWKL30ZS" Then
        DruckeStrichcodeY acArtNr(), lAnzahl, acAnzEti()
        reportbildschirmToPrinterETI "AWKL30ZS", gcEtikettenDrucker, False
    Else
        If srepname = "aWOKINE" Then
            DruckeGrundPreisEtikettenWKL30Jebe acArtNr(), lAnzahl, "NETTO"

        Else
            DruckeGrundPreisEtikettenWKL30Jebe acArtNr(), lAnzahl, "BRUTTO"
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeEtiforSofort"
    Fehler.gsFehlertext = "Im Programmteil Etiketten wählen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    Resume Next
End Sub
Private Sub SucheArtikelWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim bDebug As Boolean
    Dim iRet As Integer
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim rsrs1 As Recordset
    Dim rsRs2 As Recordset
    Dim cSuch As String
    Dim cArtNr As String
    Dim cArtBez As String
    Dim dBestand As Double
    Dim dVkPr As Double
    Dim dKVKPR As Double
    Dim dEkpr As Double
    Dim dLEKPR As Double
    Dim cLinr As String
    Dim cLiefBez As String
    Dim cLiBesNr As String
    Dim bgefunden As Boolean
    Dim cFeld As String
    Dim cLBSatz As String
    Dim lMinBest As Long
    Dim bEAN As Boolean
    Dim cEAN As String
    Dim sawm    As String
    Dim sNotizen As String
    Dim bEX As Boolean
    Dim sMWST As String
    
    Label2(9).Caption = ""
    bEX = False
    bDebug = False
    bgefunden = True
    bEAN = True
    bfoundauto = False
    
    cSuch = Text1(0).Text
    cSuch = Trim$(cSuch)
    
    If cSuch = "" Then
'        MsgBox "Bitte Wert eingeben!", vbCritical, "Winkiss Abbruch"
        Text1(0).SetFocus
        Exit Sub
    Else
        If Ist_in_ARTEAN_K(cSuch) Then
                
        End If
    End If
    
    cLinr = Text1(4).Text
    cLinr = Trim$(cLinr)
    
    If Len(cSuch) > 6 Then
        iRet = fnPruefeEANWert(cSuch)
        Select Case iRet
            Case Is = 0
                'alles okay
            Case Is = 1     'falsche Länge
                bEAN = False

            Case Is = 8     'falscher EAN-8
                bEAN = False

            Case Is = 12    'falscher UPC-A
                bEAN = False

            Case Is = 13    'falscher EAN-13
                bEAN = False

        End Select
        cSQL = "Select ARTLIEF.ARTNR, ARTIKEL.BEZEICH,ARTIKEL.MWST, ARTIKEL.AWM, ARTLIEF.LINR, ARTIKEL.BESTAND, ARTIKEL.VKPR, ARTIKEL.KVKPR1, ARTIKEL.MINBEST, ARTLIEF.LIBESNR, ARTIKEL.EAN, ARTIKEL.Notizen, ARTIKEL.RKZ  "
        cSQL = cSQL & " from ARTIKEL , ARTLIEF  where ARTIKEL.ARTNR = ARTLIEF.ARTNR "
        
'        cSQL = cSQL & " from ARTIKEL  inner join ARTLIEF  on ARTIKEL.ARTNR = ARTLIEF.ARTNR "
    End If
    
    If Len(cSuch) <= 6 Then
        cSQL = "Select ARTLIEF.ARTNR, ARTIKEL.BEZEICH,ARTIKEL.MWST, ARTIKEL.AWM, ARTLIEF.LINR, ARTIKEL.BESTAND, ARTIKEL.VKPR, ARTIKEL.KVKPR1, ARTIKEL.MINBEST, ARTLIEF.LIBESNR, ARTIKEL.EAN, ARTIKEL.Notizen, ARTIKEL.RKZ "
        cSQL = cSQL & " from ARTIKEL , ARTLIEF  where ARTIKEL.ARTNR = ARTLIEF.ARTNR  "
        cSQL = cSQL & " and ARTLIEF.ARTNR = " & cSuch & " "
'        cSQL = cSQL & " and (ARTLIEF.ARTNR = " & cSuch & " or  ARTIKEL.EAN = '" & cSuch & "' or  ARTIKEL.EAN2 = '" & cSuch & "' or  ARTIKEL.EAN3 = '" & cSuch & "') "
    Else
        If Len(cSuch) <= 8 And (Left(cSuch, 1) = "2") Then
        
            If Check7.Value = vbChecked Then
                cSQL = cSQL & " and (ARTIKEL.EAN = '" & cSuch & "' "
                cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cSuch & "' "
                cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cSuch & "' )"
            
            Else
                cSuch = Mid(cSuch, 2, 6)
                cSQL = cSQL & " and ARTLIEF.ARTNR = " & cSuch & " "
            End If
            
        ElseIf Len(cSuch) <= 8 And (Left(cSuch, 1) = "0") Then
            cSQL = cSQL & "and (ARTIKEL.EAN = '" & cSuch & "' "
            cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cSuch & "' "
            cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cSuch & "' )"
        Else
            If bEAN Then
                cSQL = cSQL & "and (ARTIKEL.EAN = '" & cSuch & "' "
                cSQL = cSQL & "or ARTIKEL.EAN2 = '" & cSuch & "' "
                cSQL = cSQL & "or ARTIKEL.EAN3 = '" & cSuch & "' )"
            Else
                cSQL = cSQL & "and ARTIKEL.LIBESNR = '" & cSuch & "' "
            End If
        End If
    End If
    
    If Len(cLinr) > 0 Then
        cSQL = cSQL & " and ARTLIEF.LINR = " & cLinr & " "
    End If
    
    cSQL = cSQL & " and ( ARTIKEL.SYNSTATUS is null or ARTIKEL.SYNSTATUS = 'E' or ARTIKEL.SYNSTATUS = 'A' ) "
    
    bgefunden = False
    
'    MsgBox cSQL
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        
        cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If rsrs.EOF Then
            
            
            If Len(cSuch) = 8 And Left(cSuch, 1) = "2" Then
                cSuch = Mid(cSuch, 2, 6)
                
                rsrs.Close: Set rsrs = Nothing
                cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
                Set rsrs = gdBase.OpenRecordset(cSQL)
                
                If rsrs.EOF Then
                
                Else
                    bgefunden = True
                End If
'                rsrs.Close: Set rsrs = Nothing
            ElseIf Len(cSuch) = 8 And Left(cSuch, 1) = "0" Then
                cSuch = Mid(cSuch, 2, 6)
                
                rsrs.Close: Set rsrs = Nothing
                cSQL = "Select * from ARTIKEL where ARTNR = " & cSuch & " "
                Set rsrs = gdBase.OpenRecordset(cSQL)
                
                If rsrs.EOF Then
                
                Else
                    bgefunden = True
                End If
'                rsrs.Close: Set rsrs = Nothing
            
            End If
        Else
            bgefunden = True
        End If
    Else
        bgefunden = True
    End If
    
    If bgefunden = True Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!artnr) Then
            cArtNr = rsrs!artnr
        Else
            cArtNr = ""
        End If
        cArtNr = Trim$(cArtNr)
        Text1(0).Text = cArtNr
        
        If Not IsNull(rsrs!BEZEICH) Then
            cArtBez = rsrs!BEZEICH
        Else
            cArtBez = ""
        End If
        cArtBez = Trim$(cArtBez)
        
        If Not IsNull(rsrs!EAN) Then
            cEAN = rsrs!EAN
        Else
            cEAN = ""
        End If
        
        If Not IsNull(rsrs!MWST) Then
            sMWST = rsrs!MWST
        Else
            sMWST = ""
        End If
        
        If Not IsNull(rsrs!RKZ) Then
            If rsrs!RKZ = "J" Then
                bEX = True
            End If
        End If
        
        If Not IsNull(rsrs!AWM) Then
            sawm = rsrs!AWM
        Else
            sawm = ""
        End If
        
        
        
        
    
        If Not IsNull(rsrs!BESTAND) Then
            dBestand = rsrs!BESTAND
        Else
            dBestand = 0
        End If
    
        If Not IsNull(rsrs!vkpr) Then
            dVkPr = rsrs!vkpr
        Else
            dVkPr = 0
        End If
    
        If Not IsNull(rsrs!KVKPR1) Then
            dKVKPR = rsrs!KVKPR1
        Else
            dKVKPR = 0
        End If
        
        If Not IsNull(rsrs!MINBEST) Then
            lMinBest = rsrs!MINBEST
        Else
            lMinBest = 0
        End If
        
        
        
        If Not IsNull(rsrs!NOTIZEN) Then
            sNotizen = rsrs!NOTIZEN
        Else
            sNotizen = ""
        End If
        sNotizen = Trim$(sNotizen)
        
        
        'Achtung
        'wenn kein Lieferant zwingend war, dann den mit dem kleinsten LEK ausser LEK = 0
        
        cLiBesNr = ""
        If cLinr <> "" Then 'linr ist vorher ausgewählt, bleibt dabei
            'clinr bleibt
            If Not IsNull(rsrs!LIBESNR) Then
                cLiBesNr = Trim(rsrs!LIBESNR)
            End If
        Else
            'clinr vom kleinsten LEK
            
            Dim rsLEK As DAO.Recordset
            Dim sEKpr As String
            Dim sEKsql As String
            Dim sSQL As String
            
            sSQL = "Select min(lekpr) as minLek from artlief where artnr = " & cArtNr
            sSQL = sSQL & " and lekpr > 0 "
            
            Set rsLEK = gdBase.OpenRecordset(sSQL)
            If Not rsLEK.EOF Then
                sEKpr = "0"
                If Not IsNull(rsLEK!minLek) Then
                    sEKpr = rsLEK!minLek
                End If
            End If
            rsLEK.Close
            
            sEKsql = SwapStr(sEKpr, ",", ".")
        
        
            sSQL = "Select linr,LIBESNR from artlief where artnr = " & cArtNr & "  and LEKPR = " & sEKsql & " "
            Set rsLEK = gdBase.OpenRecordset(sSQL)
            If Not rsLEK.EOF Then
                
                cLiBesNr = ""
                If Not IsNull(rsLEK!linr) Then
                    cLinr = rsLEK!linr
                End If
                
                If Not IsNull(rsLEK!LIBESNR) Then
                    cLiBesNr = Trim(rsLEK!LIBESNR)
                End If
            End If
            rsLEK.Close
            
        End If
        
        Text1(6).Text = cLiBesNr
    Else
        anzeige "Rot", "Artikel nicht gefunden!", Label2(0)
    End If
    rsrs.Close: Set rsrs = Nothing
    

    If bgefunden Then
    
        cLiefBez = ermLiefBez(CLng(cLinr))

        Label2(0).ForeColor = glLink
        Label2(0).Caption = cArtBez
        Label2(1).Caption = dBestand
        Label2(6).Caption = erminBestell(cArtNr)
        lbl6(43).Caption = ErmlzZugangM(cArtNr)
        
        Label1(17).Caption = ermFarbbez(sawm, cArtNr)
        Label2(2).ForeColor = ermForecolor(sawm)
        Label2(2).BackColor = ermBackcolor(sawm, cArtNr)
        Label2(2).Caption = cArtNr
        
        Label2(3).Caption = Format$(dVkPr, "##,##0.00") & " " & gcWaehrung
        Label2(5).Caption = Format$(dKVKPR, "##,##0.00") & " " & gcWaehrung
        
        Label2(10).Caption = sMWST
        
        Dim dErrechneterZentraldrogpreis As Double
        Dim cErrechneterZentraldrogpreis As String
        
        If gbBestAkt Then
            dErrechneterZentraldrogpreis = dVkPr * 80 / 100
            cErrechneterZentraldrogpreis = Runden(dErrechneterZentraldrogpreis)
            
            If CDbl(Format(cErrechneterZentraldrogpreis, "#####0.00")) <> CDbl(Format(dKVKPR, "#####0.00")) Then
                Label4(39).Caption = Format(cErrechneterZentraldrogpreis, "#####0.00")
                Label4(39).Visible = True
                Label4(39).ForeColor = vbRed
            Else
                Label4(39).Caption = ""
                Label4(39).Visible = False
            End If
        End If
        
        '**********aus ARTMERK
        Label2(8).Caption = ""
        cSQL = "Select MERK from ARTMERK where ARTNR = " & cArtNr & " "
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            If Not IsNull(rsrs1!merk) Then
                Label2(8).Caption = rsrs1!merk
            End If
        End If
        rsrs1.Close: Set rsrs1 = Nothing
        
        If bEX Then
            Label2(9).Caption = "EX"
        Else
            Label2(9).Caption = ""
        End If
        
        
        Text1(2).Text = Format$(dKVKPR, "#####0.00")
        If Trim$(Text1(4).Text) <> "" Then
            If Check8.Value = vbUnchecked Then
                Text1(4).Text = cLinr
                Label2(4).Caption = cLiefBez
            End If
        Else
            Text1(4).Text = cLinr
            Label2(4).Caption = cLiefBez
        End If
        Text1(5).Text = Format$(lMinBest, "#####0")
        Text1(0).Text = cEAN
        
        If Check5.Value = vbUnchecked Then
            Text1(13).Text = sNotizen
        End If
        
        Dim iVPE As Integer
        iVPE = ermVPE(cArtNr, cLinr)
        
        If Check6.Value = vbChecked Then 'VPE Vorschlag
    
            If iFaktor_Umverpack > 1 Then
            
                lbl6(0).Caption = iFaktor_Umverpack
                Text1(14).Text = "1"
                Text1(1).Text = 1 * iFaktor_Umverpack
                
            Else
            
                lbl6(0).Caption = iVPE
                Text1(14).Text = "1"
                Text1(1).Text = 1 * iVPE
                
            End If
        End If
        
        
        Text1(1).SetFocus
    End If
    
    If bgefunden = True Then
        LeseLieferantenPreisWKL15
    End If
    
    If bgefunden = True Then
        bfoundauto = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub

Private Function fnPruefeEingabeWKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    fnPruefeEingabeWKL15 = 1
        
    For lcount = 7 To 11
        If Trim$(Text1(lcount).Text) <> "" Then
            fnPruefeEingabeWKL15 = 0
            Exit Function
        End If
    Next lcount
               
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Check6_Click()
On Error GoTo LOKAL_ERROR

    If Check6.Value = vbChecked Then
        Text1(14).Visible = True
        Text1(14).Text = "1"
        Command7(2).Visible = True
        Command7(3).Visible = True
    Else
        Text1(14).Visible = False
        Text1(14).Text = "0"
        Command7(2).Visible = False
        Command7(3).Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check6_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Check65_Click()
On Error GoTo LOKAL_ERROR
    
    If Check65.Value = vbChecked Then
        Option1(0).Visible = True
        Option1(1).Visible = True
    Else
        Option1(0).Visible = False
        Option1(1).Visible = False
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check65_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check8_Click()
On Error GoTo LOKAL_ERROR
    
    If Check8.Value = vbUnchecked Then
        Text1(4).Text = ""
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check8_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Check9_Click()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check9.Value = vbChecked Then
        sSQL = "Update wkEINSTE Set scanmodi = True"
        gdApp.Execute sSQL, dbFailOnError
        gbscanmodi = True
        
    ElseIf Check9.Value = vbUnchecked Then
        sSQL = "Update WKEINSTE Set scanmodi = False"
        gdApp.Execute sSQL, dbFailOnError
        gbscanmodi = False
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check9_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub cmdGo_Click()
    On Error GoTo LOKAL_ERROR
    Dim iRet As Integer
        
    iRet = fnPruefeEingabeWKL15()
    If iRet <> 0 Then
        Label9.Caption = "Bitte mindestens ein Suchkriterium angeben!"
        Label9.Refresh
        Text1(8).SetFocus
        Exit Sub
    End If
    
    Label9.Caption = "Daten werden ermittelt, bitte warten..."
    Label9.Refresh
    
    Screen.MousePointer = 11

    FormatMShFlex3WKL15
    FuellenMShFlex3WKL15 SucheArtikelWKL15a



    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "cmdGo_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FuellenMShFlex3WKL15(sSQL As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim i           As Integer
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    With MSHFlex3
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            
            .Rows = lrow + 1
            .Row = lrow
            
            If Not IsNull(rsrs!artnr) Then
                lWert = rsrs!artnr
            Else
                lWert = 0
            End If
            
            .Col = 0
            .Text = lWert
            
            If Not IsNull(rsrs!BEZEICH) Then
                sWert = rsrs!BEZEICH
            Else
                sWert = ""
            End If
            
            .Col = 1
            .Text = sWert
            
            If Not IsNull(rsrs!AGN) Then
                lWert = rsrs!AGN
            Else
                lWert = "0"
            End If
            
            .Col = 2
            .Text = lWert
            
            If Not IsNull(rsrs!lekpr) Then
                siWert = rsrs!lekpr
            Else
                siWert = "0"
            End If
            
            .Col = 3
            .Text = siWert
            
            If Not IsNull(rsrs!KVKPR1) Then
                siWert = rsrs!KVKPR1
            Else
                siWert = "0"
            End If
            
            .Col = 4
            .Text = siWert
            
            If Not IsNull(rsrs!MWST) Then
                sWert = rsrs!MWST
            Else
                sWert = ""
            End If
            
            .Col = 5
            .Text = sWert
            
            If Not IsNull(rsrs!linr) Then
                lWert = rsrs!linr
            Else
                lWert = "0"
            End If
            
            .Col = 6
            .Text = lWert
            
            If Not IsNull(rsrs!LIBESNR) Then
                sWert = (rsrs!LIBESNR)
            Else
                sWert = ""
            End If
            
            .Col = 7
            .Text = sWert
            
            If Not IsNull(rsrs!EAN) Then
                sWert = rsrs!EAN
            Else
                sWert = ""
            End If
            
            .Col = 8
            .Text = sWert
            
            If Not IsNull(rsrs!EAN2) Then
                sWert = rsrs!EAN2
            Else
                sWert = ""
            End If
            
            .Col = 9
            .Text = sWert
            
            If Not IsNull(rsrs!EAN3) Then
                sWert = rsrs!EAN3
            Else
                sWert = ""
            End If
            
            .Col = 10
            .Text = sWert
            
            
            If Not IsNull(rsrs!RKZ) Then
                sWert = rsrs!RKZ
            Else
                sWert = ""
            End If
            
            .Col = 11
            .Text = sWert
            
            If Not IsNull(rsrs!LPZ) Then
                lWert = rsrs!LPZ
            Else
                lWert = "0"
            End If
            
            .Col = 12
            .Text = lWert
            
            If Not IsNull(rsrs!NOTIZEN) Then
                sWert = rsrs!NOTIZEN
            Else
                sWert = ""
            End If

            .Col = 13
            .Text = sWert
            
            If Not IsNull(rsrs!BESTAND) Then
                lWert = rsrs!BESTAND
            Else
                lWert = "0"
            End If
            
            .Col = 14
            .Text = lWert
            
            
            If Not IsNull(rsrs!GEFUEHRT) Then
                sWert = rsrs!GEFUEHRT
            Else
                sWert = ""
            End If

            .Col = 15
            .Text = sWert
            
            For i = 0 To 15
                If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
'                    aBreite(i) = Len(.TextMatrix(0, i)) * 80
                    aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                End If
            Next i
            
            rsrs.MoveNext
        Loop
    End If
    
    
    
    For i = 0 To 15
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.5
    Next i
    
    rsrs.Close: Set rsrs = Nothing
    .RowHeight(1) = 0
    lrow = lrow - 1
    
    If lrow > 1 Then
        Label9.Caption = lrow & " Artikel wurden ermittelt."
        Label9.Refresh
    ElseIf lrow = 1 Then
        Label9.Caption = lrow & " Artikel wurde ermittelt."
        Label9.Refresh
    Else
        Label9.Caption = "Es wurden keine Artikel ermittelt."
        Label9.Refresh
        Exit Sub
    End If

    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex3WKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function SucheArtikelWKL15a() As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim cFeld       As String
    Dim cwhere      As String
    Dim lcol        As Long
    Dim dWert       As Double
    Dim iRet        As Integer
    Dim cEAN        As String
    Dim cArtNr      As String
    Dim cEigNr      As String
    Dim sSQL        As String
    
    
    cSQL = "Select A.ARTNR"
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", A.LEKPR"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.MWST"
    cSQL = cSQL & ", A.LINR"
    cSQL = cSQL & ", A.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.ETIMERK"
    cSQL = cSQL & ", A.MOPREIS"
    cSQL = cSQL & ", A.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.VKMENGE"
    cSQL = cSQL & ", A.VKDATUM"
    cSQL = cSQL & ", A.MINMEN"
    cSQL = cSQL & ", A.EAN2"
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.INHALT"
    cSQL = cSQL & ", A.INHALTBEZ"
    cSQL = cSQL & ", A.GRUNDPREIS"
    cSQL = cSQL & ", A.MINBEST"
    cSQL = cSQL & ", A.RABATT_OK"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.EKPR"
    cSQL = cSQL & ", A.KVKPR1"
    
    cwhere = ""
    
    
    cSQL = cSQL & " from ARTIKEL A "
    
    
    cFeld = Text1(7).Text       'Bezeich
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.BEZEICH like '" & cFeld & "*' "
    End If
    
    cFeld = Text1(8).Text   'EAN oder ARTNR
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cEAN = cFeld
        If Len(cFeld) <= 6 Then
            cArtNr = cFeld
        Else
            cArtNr = ""
        End If
        If Left(cFeld, 1) = "2" Or Left(cFeld, 1) = "0" And Len(cFeld) = 8 Then
            cEigNr = Mid(cFeld, 2, 6)
        Else
            cEigNr = ""
        End If
        
        cwhere = cwhere & "("
        If cEAN <> "" Then
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "A.EAN like '" & cEAN & "' "
            Else
                cwhere = cwhere & "A.EAN = '" & cEAN & "' "
            End If
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "or A.EAN2 like '" & cEAN & "' "
            Else
                cwhere = cwhere & "or A.EAN2 = '" & cEAN & "' "
            End If
            If InStr(cEAN, "*") > 0 Then
                cwhere = cwhere & "or A.EAN3 like '" & cEAN & "' "
            Else
                cwhere = cwhere & "or A.EAN3 = '" & cEAN & "' "
            End If
        End If
        If cArtNr <> "" Then
            If InStr(cArtNr, "*") > 0 Then
                cwhere = cwhere & " or A.ARTNR like '" & cArtNr & "' "
            Else
                cwhere = cwhere & " or A.ARTNR = " & cArtNr & " "
            End If
        End If
        If cEigNr <> "" Then
            cwhere = cwhere & " or A.ARTNR = " & cEigNr & " "
        End If
        cwhere = cwhere & ") "
        
    End If
    
    cFeld = Text1(9).Text       'Linr
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.LINR = " & cFeld & " "
    End If
    
    cFeld = Text1(11).Text       'AGN
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.AGN = " & cFeld & " "
    End If
    
    cFeld = Text1(10).Text       'Liebesnr
    cFeld = Trim$(cFeld)
    If cFeld <> "" Then
        If cwhere = "" Then
            cwhere = "where "
        Else
            cwhere = cwhere & "and "
        End If
        cwhere = cwhere & "A.LIBESNR like '" & cFeld & "' "
    End If
    
    cSQL = cSQL & cwhere & " group by "
    cSQL = cSQL & "  A.ARTNR"
    cSQL = cSQL & ", A.BEZEICH"
    cSQL = cSQL & ", A.AGN"
    cSQL = cSQL & ", A.LEKPR"
    cSQL = cSQL & ", A.VKPR"
    cSQL = cSQL & ", A.MWST"
    cSQL = cSQL & ", A.LINR"
    cSQL = cSQL & ", A.LIBESNR"
    cSQL = cSQL & ", A.EAN"
    cSQL = cSQL & ", A.ETIMERK"
    cSQL = cSQL & ", A.MOPREIS"
    cSQL = cSQL & ", A.RKZ"
    cSQL = cSQL & ", A.LPZ"
    cSQL = cSQL & ", A.NOTIZEN"
    cSQL = cSQL & ", A.BESTAND"
    cSQL = cSQL & ", A.VKMENGE"
    cSQL = cSQL & ", A.VKDATUM"
    cSQL = cSQL & ", A.MINMEN"
    cSQL = cSQL & ", A.EAN2"
    cSQL = cSQL & ", A.EAN3"
    cSQL = cSQL & ", A.INHALT"
    cSQL = cSQL & ", A.INHALTBEZ"
    cSQL = cSQL & ", A.GRUNDPREIS"
    cSQL = cSQL & ", A.MINBEST"
    cSQL = cSQL & ", A.RABATT_OK"
    cSQL = cSQL & ", A.GEFUEHRT"
    cSQL = cSQL & ", A.EKPR"
    cSQL = cSQL & ", A.KVKPR1 "
    cSQL = cSQL & ", A.LINR "
    cSQL = cSQL & ", A.LPZ "
    cSQL = cSQL & "order by A.LINR, A.LPZ, A.BEZEICH "
    
    loeschNEW "artueb", gdBase
    
    sSQL = "Create Table ARTUEB ( "
    sSQL = sSQL & " ARTNR double"
    sSQL = sSQL & ", BEZEICH Text(35)"
    sSQL = sSQL & ", AGN double"
    sSQL = sSQL & ", LEKPR double"
    sSQL = sSQL & ", VKPR double"
    sSQL = sSQL & ", MWST Text(1)"
    sSQL = sSQL & ", LINR double"
    sSQL = sSQL & ", LIBESNR Text(13)"
    sSQL = sSQL & ", EAN Text(13)"
    sSQL = sSQL & ", ETIMERK Text(1)"
    sSQL = sSQL & ", MOPREIS double"
    sSQL = sSQL & ", RKZ Text(1)"
    sSQL = sSQL & ", LPZ double"
    sSQL = sSQL & ", NOTIZEN Text(25)"
    sSQL = sSQL & ", BESTAND double"
    sSQL = sSQL & ", VKMENGE double"
    sSQL = sSQL & ", VKDATUM DateTime"
    sSQL = sSQL & ", MINMEN double"
    sSQL = sSQL & ", EAN2 Text(13)"
    sSQL = sSQL & ", EAN3 Text(13)"
    sSQL = sSQL & ", INHALT double"
    sSQL = sSQL & ", INHALTBEZ Text(3)"
    sSQL = sSQL & ", GRUNDPREIS Text(1)"
    sSQL = sSQL & ", MINBEST double"
    sSQL = sSQL & ", RABATT_OK Text(1)"
    sSQL = sSQL & ", GEFUEHRT Text(1)"
    sSQL = sSQL & ", EKPR double"
    sSQL = sSQL & ", KVKPR1 double"
    sSQL = sSQL & " ) "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ARTUEB " & cSQL
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    SucheArtikelWKL15a = cSQL
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelWKL15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Sub FormatMShFlex3WKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim j As Integer
    
    With MSHFlex3
        .Visible = False
        .Clear
        
        .Rows = 25
        .Cols = 16
         ReDim aBreite(.Cols)
        .FixedCols = 1
        .FixedRows = 1
   
        .Row = 0
        .Col = 0
        .Text = "Artnr"
        
        .Col = 1
        .Text = "Artikelbezeichnung"
        
        .Col = 2
        .Text = "AGN"
        
        .Col = 3
        .Text = "EK - Preis"
        
        .Col = 4
        .Text = "Kassen - VK"
        
        .Col = 5
        .Text = "MWST"
       
        .Col = 6
        .Text = "Lieferant"
            
        .Col = 7
        .Text = "Lieferanten Bestnr."
              
        .Col = 8
        .Text = "EAN"
        
        .Col = 9
        .Text = "2. EAN"
        
        .Col = 10
        .Text = "3. EAN"
        
        .Col = 11
        .Text = "RKZ"
            
        .Col = 12
        .Text = "LPZ"
           
        .Col = 13
        .Text = "Notizen"
        
        .Col = 14
        .Text = "Bestand"
        
        .Col = 15
        .Text = "geführt"
        
        For j = 0 To .Cols - 1
            .Col = j
'            aBreite(j) = Len(.TextMatrix(0, j)) * 80
            aBreite(j) = TextWidth(.TextMatrix(0, j))
        Next j
        
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatMShFlex3WKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub Combo1_Change()
On Error GoTo LOKAL_ERROR


    'Achtung hier wird gerne ein EAN reingescannt

    
    If Len(Combo1.Text) > 20 Then
        Combo1.Text = Left(Combo1.Text, 20)
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_Change"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    AutocompleteCombo KeyCode, Shift, Combo1
    
    If KeyCode = vbKeyEscape Then
        Form_Unload 1
    End If
    If KeyCode = vbKeyReturn Then
            
        If Trim(Label2(0).Caption) = "unbekannt" Or (Label2(0).Caption) = "Artikel nicht gefunden!" Then
            Text1(0).SetFocus
        Else
            Text1(1).SetFocus
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 5 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Combo1_KeyUp"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub

Private Sub Combo1_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Combo1.BackColor = vbWhite
    If Combo1.Text = "" Then
        Combo1.Text = "keine"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cValid = "1234567890-," & Chr$(8)
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Combo2_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Combo2.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0
            Text3(0).Text = Format(Datumschreiben11a(3000, 9000), "DD.MM.YY")
            'fertig
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub



Private Sub Command18_Click()
On Error GoTo LOKAL_ERROR

    If Check8.Value = vbUnchecked Then
        Text1(4).Text = ""
    End If
    
    Text1(0).Text = ""
    Text1(0).SetFocus
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command18_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command1_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim cValid      As String
    Dim cFeld       As String
    Dim cZeichen    As String
    Dim lcount      As Long
    Dim bTextSuche  As Boolean
    Dim iFilnr      As Integer
    Dim cSQL        As String
    Dim rsrs        As DAO.Recordset

    Screen.MousePointer = 11
    Command6.Visible = False
    Text1(3).Text = ""
    
    If Check1.Value = vbChecked Then
        gsARTNR = ermartnrausLIBESNR(Trim$(Text1(0).Text), Val(Text1(4).Text))
        If gsARTNR <> "" Then
            Text1(0).Text = gsARTNR
            gsARTNR = ""
            bTextSuche = False
        Else
            bTextSuche = True
            gbLibesnrSeek = True
        End If
    Else
        gbLibesnrSeek = False
        cValid = "1234567890"
        cFeld = Text1(0).Text
        
        bTextSuche = False
        
        For lcount = 1 To Len(cFeld)
            cZeichen = Mid(cFeld, lcount, 1)
            If InStr(cValid, cZeichen) = 0 Then
                bTextSuche = True
                Exit For
            End If
        Next lcount
    End If
    
    If bTextSuche Then
        gcSuch = Text1(0).Text
        gsARTNR = ""
        frmWKL70.Show 1
        Me.Refresh
        If gsARTNR <> "" Then
            Text1(0).Text = gsARTNR
            gsARTNR = ""
            SucheArtikelWKL15
            gbLibesnrSeek = False
        End If
    Else
    
        '    Suche in zuordean
        
        iFaktor_Umverpack = 1
        cSQL = "Select * from Zuordean where GPEAN = '" & Text1(0).Text & "' "
        FnOpenrecordset rsrs, cSQL, 1, gdBase
        If Not rsrs.EOF Then
        
            Text1(0).Text = IIf(IsNull(rsrs!EAN), "", rsrs!EAN)
            
            iFaktor_Umverpack = IIf(IsNull(rsrs!Faktor), 1, rsrs!Faktor)
            
'            Label2(5).Caption = IIf(IsNull(rsrs!Faktor), 1, rsrs!Faktor * dAnz)
        End If
        rsrs.Close: Set rsrs = Nothing
    
        SucheArtikelWKL15
    End If
    
    If bfoundauto = False Then
    
        Dim cSuchU_EAN As String
        cSuchU_EAN = Text1(0).Text
        
        Text1(0).Text = unbekanntenEAN_Suchen_und_Anlegen(Text1(0).Text)
        
        If Text1(0).Text = "" Then
            Text1(0).Text = unbekanntenEAN_Suchen_und_Anlegen_DrogAlles(cSuchU_EAN) 'Über Drogerie und Spielwaren Schalter
        End If
        
        If Text1(0).Text <> cSuchU_EAN Then
            SucheArtikelWKL15  'also nochmal suchen
            Exit Sub
        End If
    End If
    
    If bfoundauto And fromMde = False And bscanner Then
        bscanner = False
        bfoundauto = False
        Command2_Click 15
    Else
        fromMde = False
    End If

    iFilnr = CInt(gcFilNr)
    If iFilnr > 0 And Label2(2).Caption <> "" Then
        Command6.Visible = True
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command11_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 0, 3, 4
            bLieferavis = False
            Unload frmWKL15
        Case Is = 1
'            bLieferavis = False
            Frame4.Visible = False
            
        Case Is = 2
            Frame6.Visible = False
            Frame7.Visible = True

        Case Is = 5
            Frame7.Visible = False
            anzeigenalleDS
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command12_Click()
    On Error GoTo LOKAL_ERROR
    
    gsZSpalte = "EAN"
    gstab = "WAEINGEM"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command12_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command13_Click()
    On Error GoTo LOKAL_ERROR
    
    If glSelect < 1 Then
            MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    Else
        If Unbekannt(glSelect - 1) Then
            Frame7.Visible = True
        Else
            Label4(0).Caption = "Sie können nur unbekannte EAN - Codes zu bekannten Artikeln zuordnen."
            Label4(0).Refresh
        End If
    End If
            
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function Unbekannt(lauswahl As Long) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim rsWA        As Recordset
    Dim sSQL        As String
    Dim sSTATUS     As String
   
    Unbekannt = False
    
    sSQL = "Select * from Waeingem where Reihenf = " & lauswahl
    sSQL = sSQL & " and Status = 'unbekannt' "
    Set rsWA = gdBase.OpenRecordset(sSQL)
    If Not rsWA.RecordCount = 0 Then
        Unbekannt = True
    End If
    rsWA.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Unbekannt"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub WAEINGEMUpdaten()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    sSQL = " Update WAEINGEM inner join artikel on waeingem.ean = artikel.ean "
    sSQL = sSQL & " set waeingem.artnr = artikel.artnr , waeingem.bezeich = artikel.bezeich"
    sSQL = sSQL & " , waeingem.status = 'bekannt'"
    sSQL = sSQL & " , waeingem.agn = artikel.agn "
    sSQL = sSQL & " , waeingem.LINR = artikel.LINR "
    sSQL = sSQL & " , waeingem.LPZ = artikel.LPZ "
    sSQL = sSQL & " , waeingem.LIBESNR = artikel.LIBESNR "
    sSQL = sSQL & " , waeingem.EAN = artikel.EAN "
    sSQL = sSQL & " , waeingem.RKZ = artikel.RKZ "
    sSQL = sSQL & " , waeingem.NOTIZEN = artikel.NOTIZEN "
    sSQL = sSQL & " , waeingem.ETIMERK = artikel.ETIMERK "
    sSQL = sSQL & " , waeingem.LEK = artikel.LEKPR "
    sSQL = sSQL & " , waeingem.NSN = artikel.SPANNE "
    sSQL = sSQL & " , waeingem.KVKPR = artikel.KVKPR1 "
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WAEINGEMUpdaten"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WAEINGEMUpdatenEinzeln(lreihenfolge As Long, statement As String)
    On Error GoTo LOKAL_ERROR
    
    Dim rsArt       As Recordset
    Dim rsWA        As Recordset
    Dim sSQL        As String
    Dim sSTATUS     As String
    Dim rsArtlief   As Recordset
    
    Set rsArt = gdBase.OpenRecordset(statement)
    If Not rsArt.RecordCount = 0 Then
        sSQL = "Select * from Waeingem where Reihenf = " & lreihenfolge
        Set rsWA = gdBase.OpenRecordset(sSQL)
        If Not rsWA.RecordCount = 0 Then
            sSTATUS = "bekannt"
            rsWA.Edit
            rsWA!artnr = rsArt!artnr
            rsWA!BEZEICH = rsArt!BEZEICH
            rsWA!AGN = rsArt!AGN
            rsWA!linr = rsArt!linr
            rsWA!LPZ = rsArt!LPZ
            rsWA!LIBESNR = rsArt!LIBESNR
            rsWA!RKZ = rsArt!RKZ
            rsWA!NOTIZEN = rsArt!NOTIZEN
            rsWA!ETIMERK = rsArt!ETIMERK
            rsWA!LEK = rsArt!lekpr
            rsWA!nsn = rsArt!SPANNE
            rsWA!kvkpr = rsArt!KVKPR1
            
            rsWA!Status = sSTATUS
  
            rsWA.Update
            
            'neue EAN in die Artikel eintragen
            rsArt.Edit
            
            If IsNull(rsArt!EAN2) Then
                rsArt!EAN2 = rsWA!EAN
            Else
                rsArt!EAN3 = rsWA!EAN
            End If

            rsArt.Update
            
            'neue Kombination in die Artlief eintragen
            sSQL = "Select * from Artlief where artnr = " & rsArt!artnr
            sSQL = sSQL & " and LINR = " & rsArt!linr
            Set rsArtlief = gdBase.OpenRecordset(sSQL)
            
            If rsArtlief.RecordCount = 0 Then
                rsArtlief.AddNew
                rsArtlief!SYNStatus = "A"
            Else
                rsArtlief.Edit
                rsArtlief!SYNStatus = "E"
            End If
            
            rsArtlief!artnr = rsArt!artnr
            rsArtlief!linr = rsArt!linr
            rsArtlief!lekpr = rsArt!lekpr
            rsArtlief!MINMEN = rsArt!MINMEN
            rsArtlief!LIBESNR = rsArt!LIBESNR
            
            rsArtlief.Update
            rsArtlief.Close: Set rsArtlief = Nothing
        End If
        rsWA.Close
    End If
    rsArt.Close: Set rsArt = Nothing
    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WAEINGEMUpdatenEinzeln"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Artikeluebernahme()
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Long
    Dim j           As Long
    Dim lcol        As Long
    Dim lRows       As Long
    Dim rsArt       As Recordset
    Dim rsArtlief   As Recordset
    Dim sSQL        As String
    Dim Editartikel As ArtikelTyp
    
    
    'Artikel in der Tabelle Artikel anlegen

    For i = 2 To MSHFlex2.Rows - 1
        
        lcol = 5
        MSHFlex2.Row = i
        MSHFlex2.Col = lcol
        
        If MSHFlex2.Text <> "" Then
            Editartikel.artnr = CDbl(MSHFlex2.Text)
        Else
            Editartikel.artnr = 0
        End If
        
        sSQL = "Select * from Artikel where artnr = " & Editartikel.artnr
        Set rsArt = gdBase.OpenRecordset(sSQL)
        
        MSHFlex2.Col = 2
        Editartikel.EAN = IIf(IsNull(MSHFlex2.Text), "", MSHFlex2.Text)
        MSHFlex2.Col = 3
        Editartikel.EAN2 = IIf(IsNull(MSHFlex2.Text), "", MSHFlex2.Text)
        MSHFlex2.Col = 4
        Editartikel.EAN3 = IIf(IsNull(MSHFlex2.Text), "", MSHFlex2.Text)
        
        
        MSHFlex2.Col = 6 'Pflichtfeld
        Editartikel.BEZEICH = IIf(IsNull(MSHFlex2.Text), Null, MSHFlex2.Text)
        
        MSHFlex2.Col = 7 'Pflichtfeld
        If MSHFlex2.Text <> "" Then
            Editartikel.linr = CDbl(MSHFlex2.Text)
        Else
            Editartikel.linr = 0
        End If

        MSHFlex2.Col = 8
        If MSHFlex2.Text <> "" Then
            Editartikel.LPZ = CDbl(MSHFlex2.Text)
        Else
            Editartikel.LPZ = 0
        End If
        
        MSHFlex2.Col = 9
        Editartikel.LIBESNR = IIf(IsNull(MSHFlex2.Text), "", MSHFlex2.Text)
        
        MSHFlex2.Col = 10
        If MSHFlex2.Text <> "" Then
            Editartikel.lekpr = CDbl(MSHFlex2.Text)
        Else
            Editartikel.lekpr = 0
        End If
        
        MSHFlex2.Col = 11
        If MSHFlex2.Text <> "" Then
            Editartikel.KVKPR1 = CDbl(MSHFlex2.Text)
        Else
            Editartikel.KVKPR1 = 0
        End If
        
        MSHFlex2.Col = 12
        If MSHFlex2.Text <> "" Then
            Editartikel.AGN = CDbl(MSHFlex2.Text)
        Else
            Editartikel.AGN = 0
        End If
        
        MSHFlex2.Col = 13
        Editartikel.MWST = IIf(IsNull(MSHFlex2.Text), "V", MSHFlex2.Text)
        
        Editartikel.AUFDAT = DateValue(Now)
        Editartikel.AUFSCHLAG = 0
        Editartikel.AWM = "0"
        Editartikel.BESTAND = 0
        Editartikel.BONUS_OK = "J"
        Editartikel.ekpr = 0
        Editartikel.ETIMERK = "N"
        Editartikel.GEFUEHRT = "J"
        Editartikel.GRUNDPREIS = "N"
        Editartikel.INHALT = 0
        Editartikel.INHALTBEZ = ""
        Editartikel.LASTDATE = DateValue(Now)
        Editartikel.LASTTIME = ""
        Editartikel.MINBEST = 0
        Editartikel.MINMEN = 0
        Editartikel.MOPREIS = 0
        Editartikel.NOTIZEN = ""
        Editartikel.PREISSCHU = "J"
        Editartikel.RABATT_OK = "J"
        Editartikel.RKZ = "N"
        Editartikel.SPANNE = 0
        Editartikel.UMS_OK = "J"
        Editartikel.vkpr = 0

        rsArt.Edit
        
        rsArt!artnr = Editartikel.artnr
        rsArt!AGN = Editartikel.AGN
        rsArt!AUFDAT = Editartikel.AUFDAT
        rsArt!AUFSCHLAG = Editartikel.AUFSCHLAG
        rsArt!AWM = Editartikel.AWM
        rsArt!BESTAND = Editartikel.BESTAND
        rsArt!BEZEICH = Editartikel.BEZEICH
        rsArt!BONUS_OK = Editartikel.BONUS_OK
        rsArt!EAN = Editartikel.EAN
        rsArt!EAN2 = Editartikel.EAN2
        rsArt!EAN3 = Editartikel.EAN3
        rsArt!ekpr = Editartikel.ekpr
        rsArt!ETIMERK = Editartikel.ETIMERK
        rsArt!GEFUEHRT = Editartikel.GEFUEHRT
        rsArt!GRUNDPREIS = Editartikel.GRUNDPREIS
        rsArt!INHALT = Editartikel.INHALT
        rsArt!INHALTBEZ = Editartikel.INHALTBEZ
        rsArt!KVKPR1 = Editartikel.KVKPR1
        rsArt!LASTDATE = Editartikel.LASTDATE
        rsArt!LASTTIME = Editartikel.LASTTIME
        rsArt!lekpr = Editartikel.lekpr
        rsArt!LIBESNR = Editartikel.LIBESNR
        rsArt!linr = Editartikel.linr
        rsArt!LPZ = Editartikel.LPZ
        rsArt!MINBEST = Editartikel.MINBEST
        rsArt!MINMEN = Editartikel.MINMEN
        rsArt!MOPREIS = Editartikel.MOPREIS
        rsArt!MWST = Editartikel.MWST
        rsArt!NOTIZEN = Editartikel.NOTIZEN
        rsArt!PREISSCHU = Editartikel.PREISSCHU
        rsArt!RABATT_OK = Editartikel.RABATT_OK
        rsArt!RKZ = Editartikel.RKZ
        rsArt!SPANNE = Editartikel.SPANNE
        rsArt!UMS_OK = Editartikel.UMS_OK
        rsArt!vkpr = Editartikel.vkpr
    
        rsArt.Update
        rsArt.Close: Set rsArt = Nothing
        
        lcol = 5
        'Artikel in der Artlief anlegen
    
        sSQL = "Select * from Artlief where artnr = " & Editartikel.artnr
        sSQL = sSQL & " and LINR = " & Editartikel.linr
        Set rsArtlief = gdBase.OpenRecordset(sSQL)
        If rsArtlief.RecordCount = 0 Then
            rsArtlief.AddNew
            rsArtlief!SYNStatus = "A"
        Else
            rsArtlief.Edit
            rsArtlief!SYNStatus = "E"
        End If
        rsArtlief!artnr = Editartikel.artnr
        rsArtlief!linr = Editartikel.linr
        rsArtlief!lekpr = Editartikel.lekpr
        rsArtlief!MINMEN = Editartikel.MINMEN
        rsArtlief!LIBESNR = Editartikel.LIBESNR
        
        rsArtlief.Update
        rsArtlief.Close: Set rsArtlief = Nothing
    
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Artikeluebernahme"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command14_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Artikeluebernahme
    WAEINGEMUpdaten
    
    sSQL = "Delete from Artikel where bezeich is null or bezeich = ' ' or bezeich ='' "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeigenalleDS
    
    Frame6.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command14_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command15_Click()
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    Dim sArtnr As String
    
    If lSelect < 1 Then
            MsgBox "Bitte einen Satz in der Tabelle markieren!", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    Else
        MSHFlex3.Col = 0
        MSHFlex3.Row = lSelect
        sArtnr = MSHFlex3.Text
        sSQL = "Select * from artikel where artnr = " & sArtnr
        
        WAEINGEMUpdatenEinzeln glSelect - 1, sSQL
        anzeigenalleDS
        Frame7.Visible = False
        Frame4.Visible = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command15_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command16_Click()
    On Error GoTo LOKAL_ERROR
    
    Frame7.Visible = False
    Frame6.Visible = True
    FormatMShFlex2WKL15a     'Tabelle erstellen
    FuellenMShFlex2WKL15a    'Tabelle füllen
    
     Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command16_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim lcount As Long
    Dim ctmp As String
    Dim iRet As Integer
    lcount = Val(Label3.Caption)
    
    Select Case Index
        Case 0 To 10
            If lcount >= 0 Then
                Text1(lcount).Text = Text1(lcount).Text & Command2(Index).Caption
                Text1(lcount).SetFocus
                Text1(lcount).SelLength = Len(Text1(lcount).Text)
            End If
            
        Case Is = 11        '** Plus-Zeichen **
            If lcount = 1 Then
                If InStr(1, Text1(lcount).Text, "+") > 0 Then
                    Exit Sub
                ElseIf InStr(1, Text1(lcount).Text, "-") > 0 Then
                    ctmp = Text1(lcount).Text
                    Mid(ctmp, 1, 1) = "+"
                    Text1(lcount).Text = ctmp
                Else
                    Text1(lcount).Text = "+"
                End If
            End If
            Text1(lcount).SetFocus
            
        Case Is = 12        '** Minus-Zeichen **
            If lcount = 1 Then
                If InStr(1, Text1(lcount).Text, "-") > 0 Then
                    Exit Sub
                ElseIf InStr(1, Text1(lcount).Text, "+") > 0 Then
                    ctmp = Text1(lcount).Text
                    Mid(ctmp, 1, 1) = "-"
                    Text1(lcount).Text = ctmp
                Else
                    Text1(lcount).Text = "-"
                End If
            End If
            Text1(lcount).SetFocus
        Case Is = 13        '** Löschen **
            Text1(lcount).Text = ""
            Text1(lcount).SetFocus
            
        Case Is = 14        '** Rückgängig **
            If Len(Text1(lcount).Text) > 0 Then
                ctmp = Text1(lcount).Text
                ctmp = Left(ctmp, Len(ctmp) - 1)
                Text1(lcount).Text = ctmp
            End If
            Text1(lcount).SetFocus
            
        Case Is = 15        '** Speichern **
        
            voreinstellungspeichernE15B
            
            'Überprüfe eventuell das MDH - Datum
            If Text3(0).Text <> "" Then
                If IsDate(Text3(0).Text) = False Then
                    Screen.MousePointer = 0
                    MsgBox "Bitte geben Sie ein Datum an! ('TT.MM.JJ')", vbInformation, "Winkiss Hinweis:"
                    Text3(0).SetFocus
                    Exit Sub
                End If
            End If
            
            If Trim$(Text1(1).Text) = "" Then
                Text1(1).Text = "0"
            End If
            
            If Trim$(Text1(0).Text) = "" Then
                If Label2(2).Caption = "0" Then
                    Screen.MousePointer = 0
                    MsgBox "Bitte einen Artikel festlegen!", vbInformation, "Winkiss Hinweis:"
                    Text1(0).SetFocus
                    Exit Sub
                Else
                    Text1(0).Text = Label2(2).Caption
                End If
            End If
            
            glBestandNeu = Val(Text1(1).Text)
            If glBestandNeu < 0 Then
                If glLevel < 7 Then
                    MsgBox "Mengen-Reduzierung nicht möglich!" & vbCrLf & vbCrLf & "Bestandsminderungen sind nur mit Zugriffs-Level 7 oder höher erlaubt!", vbInformation, "Winkiss Hinweis:"
                    Exit Sub
                End If
            End If
            
            If Trim$(Combo1.Text) = "" Then 'Frage nach LS Nummer
                iRet = MsgBox("Möchten Sie eine Lieferschein-Nr. eingeben?", vbCritical + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                If iRet = vbYes Then
                    Combo1.SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
                ElseIf iRet = vbNo Then
                    Combo1.Text = "keine"
                    Text1(1).SetFocus
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
            
            If Label2(7).Caption = "Stück" Then
                SchreibeDatenWKL15
            ElseIf Label2(7).Caption = "Kg" Then
                SchreibeDatenWKL15fürKilo
            End If
            
        Case Is = 16        'Vorheriges Feld
            If lcount > 0 Then
                Text1(0).SetFocus
            Else
                Text1(lcount).SetFocus
            End If
            
        Case Is = 17        'Nächstes Feld
            If lcount < 1 Then
                Text1(1).SetFocus
            Else
                Text1(lcount).SetFocus
            End If
            
        Case Is = 18        'Komma
            If lcount = 2 Or lcount = 3 Then
                If InStr(Text1(lcount).Text, ",") = 0 Then
                    Text1(lcount).Text = Text1(lcount).Text & Command2(Index).Caption
                End If
                Text1(lcount).SetFocus
                Text1(lcount).SelLength = Len(Text1(lcount).Text)
            End If
        Case Is = 19        'F2
            Text1_KeyUp Val(Label3.Caption), vbKeyF2, 0
            
        Case Is = 20        'F4
            If Text1(0).Text = "" Then
                MsgBox "Bitte den Artikel eindeutig definieren (Artikelnummer oder EAN-Code)!", vbCritical, "STOP!"
                Text1(0).SetFocus
                Exit Sub
            Else
                Text1_KeyUp Val(Label3.Caption), vbKeyF4, 0
            End If
        Case Is = 21    'Tastatur ein/aus
            If Frame1.Visible Then
                Frame1.Visible = False
            Else
                Frame1.Visible = True
            End If
        Case Is = 22 'Kundenbestellungen anzeigen
            KB "GELIEFERT", "INFORMIEREN"
            UpdateKuBestKUNDENSTATUS "INFORMIEREN", "GELIEFERT"
            Command2(22).Visible = False
        
    End Select
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Form_Unload 1
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeListe()
    On Error GoTo LOKAL_ERROR
    
    Dim sTempstring As String
    
    
    sTempstring = Space(6 - Len(CStr(newArtikel.artnr))) & newArtikel.artnr & " "
    
    sTempstring = sTempstring & Space(35 - Len(newArtikel.BEZEICH)) & newArtikel.BEZEICH & " "
    
    sTempstring = sTempstring & Space(4 - Len(CStr(newArtikel.ZubuchMe))) & newArtikel.ZubuchMe & "  "
    
    sTempstring = sTempstring & Space(8 - Len(Format$(newArtikel.REKPR, "####0.00"))) & Format$(newArtikel.REKPR, "####0.00") & " "
    
    sTempstring = sTempstring & Space(8 - Len(Format$(newArtikel.lekpr, "####0.00"))) & Format$(newArtikel.lekpr, "####0.00") & " "
    
    sTempstring = sTempstring & Space(8 - Len(Format$(newArtikel.ekpr, "####0.00"))) & Format$(newArtikel.ekpr, "####0.00") & " "
    
    sTempstring = sTempstring & Space(8 - Len(Format$(newArtikel.KVKPR1, "####0.00"))) & Format$(newArtikel.KVKPR1, "####0.00")
    
    
    List5.AddItem sTempstring, 0
    
    
    Label11(9).Caption = Label11(9).Caption + (newArtikel.REKPR * newArtikel.ZubuchMe)
    Label11(9).Caption = Format$(Label11(9).Caption, "####0.00")
    Label11(10).Caption = Label11(10).Caption + (newArtikel.KVKPR1 * newArtikel.ZubuchMe)
    Label11(10).Caption = Format$(Label11(10).Caption, "####0.00")
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeListe"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SchreibeListeKg()
    On Error GoTo LOKAL_ERROR
    
    Dim sTempstring As String
   
    sTempstring = Space(6 - Len(CStr(newArtikel.artnr))) & newArtikel.artnr & " "
    
    sTempstring = sTempstring & Space(35 - Len(newArtikel.BEZEICH)) & newArtikel.BEZEICH & " "
    
    sTempstring = sTempstring & Space(9 - Len(Format$(newArtikel.ZubuchKg, "####0.000"))) & Format$(newArtikel.ZubuchKg, "####0.000")
    
    List5.AddItem sTempstring, 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeListeKg"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Check65.Value = vbChecked Then
        sSQL = "Update DBEINSTE Set ETIKVKAE = True"
        gdBase.Execute sSQL, dbFailOnError
        gbETIKVKAE = True
        
    ElseIf Check65.Value = vbUnchecked Then
        sSQL = "Update DBEINSTE Set ETIKVKAE = False"
        gdBase.Execute sSQL, dbFailOnError
        gbETIKVKAE = False
        
    End If
    
    Unload frmWKL15
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command4_Click()
On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 0
        
    gsBackcolor = Label4(32).BackColor
    gsForecolor = Label4(32).ForeColor
    gsArtikelFarbe = Label4(32).Tag
    
    frmWKL49.Show 1
            
    Label4(32).BackColor = gsBackcolor
    Label4(32).ForeColor = gsForecolor
    Label4(32).Tag = gsArtikelFarbe
    
    If gsArtikelFarbe <> "" Then
        Label4(32).Caption = "Farbauswahl"
    Else
        Label4(32).Caption = "alle Farben"
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command5_Click()
On Error GoTo LOKAL_ERROR

Dim lLoeschen As Long
Dim cSQL As String

Dim cPfad2      As String
   
cPfad2 = gcDBPfad
If Right(cPfad2, 1) <> "\" Then
    cPfad2 = cPfad2 & "\"
End If


If tableSuchenDBKombi("ZuTemp", 2) Then
    lLoeschen = MsgBox("Druckdaten nach dem Drucken löschen?", vbQuestion + vbYesNo, "Winkiss Frage:")
    
    If Modul6.FindFile(App.Path, "aWKL15s.rpt") Then
        reportbildschirmApp "dWKL15", "aWKL15s"
    Else
    
    
    
        'LAGERP
        If Not SpalteInTabellegefundenNEW("ZuTemp", "LAGERP", gdApp) Then
            SpalteAnfuegenNEW "ZuTemp", "LAGERP", "long", gdApp
        End If
        
        'LEKPR
        If Not SpalteInTabellegefundenNEW("ZuTemp", "LEKPR", gdApp) Then
            SpalteAnfuegenNEW "ZuTemp", "LEKPR", "double", gdApp
        End If

        'Filiale
        If Not SpalteInTabellegefundenNEW("ZuTemp", "Filiale", gdApp) Then
            SpalteAnfuegenNEW "ZuTemp", "Filiale", "integer", gdApp
        End If

        'Filialbezeichnung
        If Not SpalteInTabellegefundenNEW("ZuTemp", "Filialbezeichnung", gdApp) Then
            SpalteAnfuegenNEW "ZuTemp", "Filialbezeichnung", "Text(50)", gdApp
        End If
        
        If gcFilNr <> "0" Then
            cSQL = "Update ZuTemp set Filiale = " & gcFilNr
            gdApp.Execute cSQL, dbFailOnError
            
            Dim sFilbez As String
            sFilbez = ermFilBez(CInt(gcFilNr))
            
            cSQL = "Update ZuTemp set Filialbezeichnung = '" & sFilbez & "'"
            gdApp.Execute cSQL, dbFailOnError
        End If
        
        cSQL = "Update ZuTemp set LEKPR = 0 "
        gdApp.Execute cSQL, dbFailOnError
        
        
        cSQL = "UPDATE ZuTemp AS A INNER JOIN"
        cSQL = cSQL & "[;DATABASE=" & cPfad2 & "kissdata.mdb;pwd=" & gsPasswort & "].LAGERPLATZ AS B ON A.artnr = b.artnr"
        cSQL = cSQL & " SET A.LAGERP = B.LAGERP "
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "UPDATE ZuTemp AS A INNER JOIN"
        cSQL = cSQL & "[;DATABASE=" & cPfad2 & "kissdata.mdb;pwd=" & gsPasswort & "].Artlief AS B ON A.artnr = b.artnr and A.linr = b.linr"
        cSQL = cSQL & " SET A.LEKPR = B.LEKPR "
        gdApp.Execute cSQL, dbFailOnError
    
       
        cSQL = "Update ZuTemp set LAGERP = 0 where lagerp is null "
        gdApp.Execute cSQL, dbFailOnError
    
        reportbildschirmApp "dWKL15", "aWKL15"
        
        
    End If
Else
    MsgBox "Es sind keine Druckdaten vorhanden.", vbInformation, "Winkiss Hinweis:"
End If
    
If lLoeschen = vbYes Then
    loeschapp "zutemp"
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command5_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    If KeyCode = vbKeyEscape Then
        Form_Unload 1
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Command6_Click()
    On Error GoTo LOKAL_ERROR
    
    
    
    gcArtNrFiliale = Trim(Label2(2).Caption)
    frmWKLae.Show 1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Command7_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case 0
            If Label2(2).Caption <> "0" Then
                Screen.MousePointer = 11
                gsARTNR = Label2(2).Caption
                If gsARTNR <> "" Then
                    frmWKL10.Show 1
                    Me.Refresh
                End If
                gsARTNR = ""
                Screen.MousePointer = 0
        
            Else
                Screen.MousePointer = 0
                MsgBox "Bitte einen Artikel festlegen!", vbInformation, "Winkiss Hinweis:"
                Text1(0).SetFocus
            End If
        Case 1
            Screen.MousePointer = 11
            DESADV_anwenden
        Case 2  'rauf
            Text1(14).Text = Val(Text1(14).Text) + 1
            Text1(1).Text = Val(lbl6(0).Caption) * Val(Text1(14).Text)
            Text1(1).SetFocus
            
        Case 3  'runter
            Text1(14).Text = Val(Text1(14).Text) - 1
            Text1(1).Text = Val(lbl6(0).Caption) * Val(Text1(14).Text)
            Text1(1).SetFocus
        Case 4
            
            Frame2.Visible = False
        Case 5
            del_DESADV
        Case 6
        
            bLieferavis = True
            If Excelimport = True Then
                Frame2.Visible = False
                Frame4.Visible = True
                
                anzeigenalleDS
            End If
            
        Case 7
            MsgBox "Hier wird eine Excel Tabelle erwartet:", vbInformation, "Winkiss Hinweis:"
           
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function pfadseekExcel_DESADV() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    Dim sExcelpfad  As String
    
    pfadseekExcel_DESADV = False

    sTitle = "Speichern des Pfades"
    
    sFilter = "Excel - Dateien (*.xls)|*.xls|Tchibo - Dateien (*.txt)|*.txt"
    
    
    
    sOldpfad = gcDBPfad & "\IN"
    sExcelpfad = pfadaendernKomplett(sTitle, sFilter, sOldpfad)
    
    If UCase(Right(sExcelpfad, 4)) = ".XLS" Then
        pfadseekExcel_DESADV = True
        Label18.Caption = sExcelpfad
    ElseIf UCase(Right(sExcelpfad, 4)) = ".TXT" Then
        pfadseekExcel_DESADV = True
        Label18.Caption = sExcelpfad
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcel_DESADV"
    Fehler.gsFehlertext = "Im Programmteil Terminpreise ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function Excelimport() As Boolean
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim dbExcel     As Database
    Dim rsrs        As Recordset
    Dim rsRs2       As Recordset
    Dim gsExcel50   As String
    Dim rsEAN       As Recordset
    Dim sArtnr      As String
    Dim sSTATUS     As String
    Dim i           As Integer
    
    i = 0
    
    Excelimport = False
    
    gsExcel50 = "Excel 5.0;"
    
    If pfadseekExcel_DESADV = False Then
        anzeige "rot2", "Abbruch durch Benutzer", Label17
        Exit Function
    End If
    
    anzeige "normal", "", Label17
    cPfad = Label18.Caption
    
    If UCase(Right(cPfad, 4)) = ".XLS" Then
    
        Set dbExcel = OpenDatabase(cPfad, 0, 0, gsExcel50)
    
        loeschNEW "WAEINGEM", gdBase
        CreateTableT2 "WAEINGEM", gdBase
    
        sSQL = "Select * from WAEINGEM"
        Set rsRs2 = gdBase.OpenRecordset(sSQL)
        
        Dim lLinr As Long
        lLinr = checkLinrForKISS(Label3)
        
        If lLinr = 0 Then
            Screen.MousePointer = 0
            anzeige "rot", "Keine auswertbaren Lieferantennummern enthalten.", Label17
            Exit Function
        End If
        
        Screen.MousePointer = 11
    
        Dim slibesnr As String
        Dim sVPEMENGE As String
        Dim iCount As Integer
    
        Dim rsArtlief As DAO.Recordset
        Set rsrs = dbExcel.OpenRecordset("WE$")
        If Not rsrs.EOF Then
            rsrs.MoveLast
            iCount = rsrs.RecordCount
            
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                anzeige "normal", "noch " & iCount, Label17
                iCount = iCount - 1
            
                i = i + 1
                If Not IsNull(rsrs!LIBESNR) Then
                    slibesnr = rsrs!LIBESNR
                    
                    sVPEMENGE = rsrs!VPEMENGE
                    sVPEMENGE = SwapStr(sVPEMENGE, "000", "")
                    
                    While Len(Trim(slibesnr)) < 11
                        slibesnr = "0" & slibesnr
                    Wend
                   
                    sSQL = "select * from artlief where left(libesnr,11) = '" & slibesnr & "'"
                    sSQL = sSQL & " and linr = " & lLinr
                    Set rsArtlief = gdBase.OpenRecordset(sSQL)
                    If Not rsArtlief.EOF Then
                    
                        If Not IsNull(rsArtlief!artnr) Then
                            sArtnr = rsArtlief!artnr
                        Else
                            sArtnr = ""
                        End If
                        
                        If sArtnr <> "" Then
                
                            sSQL = "select * from artikel where artnr = " & sArtnr
                            Set rsEAN = gdBase.OpenRecordset(sSQL)
                            
                            rsRs2.AddNew
                            If Not rsEAN.EOF Then
                                sSTATUS = "bekannt"
                        
                                rsRs2!artnr = rsEAN!artnr
                                rsRs2!BEZEICH = rsEAN!BEZEICH
                                rsRs2!EAN = rsEAN!EAN
                                rsRs2!Reihenf = i
                                rsRs2!Status = sSTATUS
                                rsRs2!AGN = rsEAN!AGN
                                rsRs2!linr = rsArtlief!linr
                                rsRs2!LPZ = rsEAN!LPZ
                                rsRs2!LIBESNR = rsArtlief!LIBESNR
                                rsRs2!RKZ = rsEAN!RKZ
                                rsRs2!NOTIZEN = rsEAN!NOTIZEN
                                rsRs2!BESTAND = rsEAN!BESTAND
                                rsRs2!kvkpr = rsEAN!KVKPR1
                                rsRs2!LEK = rsArtlief!lekpr
                                rsRs2!Menge = CLng(sVPEMENGE) * rsArtlief!MINMEN
                                rsRs2!ETIMERK = rsEAN!ETIMERK
                                rsRs2!nsn = rsEAN!SPANNE
                             Else
                                sSTATUS = "unbekannt"
                                rsRs2!LIBESNR = rsrs!LIBESNR
                                rsRs2!Menge = sVPEMENGE
                                rsRs2!Reihenf = i
                                rsRs2!Status = sSTATUS
                            End If
                            
                            rsRs2.Update
                            
                            rsEAN.Close
                        Else
                        
                            rsRs2.AddNew
                            sSTATUS = "unbekannt"
                                
                            rsRs2!LIBESNR = rsrs!LIBESNR
                            rsRs2!Menge = sVPEMENGE
                            rsRs2!Reihenf = i
                            rsRs2!Status = sSTATUS
                            
                            rsRs2.Update
                        End If
                    Else
                        
                        rsRs2.AddNew
                        sSTATUS = "unbekannt"
                            
                        rsRs2!LIBESNR = rsrs!LIBESNR
                        rsRs2!Menge = sVPEMENGE
                        rsRs2!Reihenf = i
                        rsRs2!Status = sSTATUS
                        
                        rsRs2.Update
                    
                    End If
                    rsArtlief.Close: Set rsArtlief = Nothing
                End If
                    
            rsrs.MoveNext
            Loop
            
        End If
        
        rsrs.Close: Set rsrs = Nothing
        
        Screen.MousePointer = 0
    
        dbExcel.Close
        
        Excelimport = True
    
    ElseIf UCase(Right(cPfad, 4)) = ".TXT" Then
    
        Dim lPosEnde        As Long
        Dim cEinzelsatz     As String
        Dim lLenfil         As Long
        Dim lposSemi        As Long
        Dim lposSemiEnde    As Long
        Dim cWert           As String
        Dim cSCAN           As String
        Dim cLiBesNr        As String
        Dim cMenge          As String
        Dim lPos            As String
        Dim cSatz1          As String
        Dim lcount          As Long
        Dim iFileNr         As Integer
        
        loeschNEW "WAEINGEM", gdBase
        CreateTableT2 "WAEINGEM", gdBase
    
        sSQL = "Select * from WAEINGEM"
        Set rsRs2 = gdBase.OpenRecordset(sSQL)
        
        
        lLinr = checkLinrForKISS(Label3)
        
        If lLinr = 0 Then
            Screen.MousePointer = 0
            anzeige "rot", "Keine auswertbaren Lieferantennummern enthalten.", Label17
            Exit Function
        End If
        
        If FileExists(cPfad) = False Then
            Exit Function
        End If
        
        
        Screen.MousePointer = 11
        
        lPos = 1
        lPosEnde = 1
        lposSemiEnde = 1
        
        iFileNr = FreeFile
        Open cPfad For Binary As #iFileNr
        If LOF(iFileNr) > 0 Then
        

            cSatz1 = Space$(LOF(iFileNr))
            Get #iFileNr, 1, cSatz1
        
            lLenfil = Len(cSatz1)
            
            lPosEnde = InStr(lPos, cSatz1, vbCr)
    
            lcount = 0
            
            Do
                lcount = lcount + 1
                lPosEnde = InStr(lPos, cSatz1, vbCr)
                
                If lPosEnde > 0 Then
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde - lPos)
                    
                    If cEinzelsatz <> "" Then
                        lPos = lPos + lPosEnde - lPos + 2
                        
                        lposSemi = 1
                        
                        cSCAN = ""
                        cMenge = ""
                        cLiBesNr = ""
                        
                        For i = 1 To 8
                            lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";")
                            cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                            lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        Next i
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cSCAN = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cLiBesNr = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, " ")
                        cMenge = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)

                        If cLiBesNr <> "" Then
                            
                            slibesnr = cLiBesNr
                            sVPEMENGE = cMenge
                            sVPEMENGE = SwapStr(sVPEMENGE, ".", ",")
                            
                            sSQL = "select * from artlief where libesnr = '" & slibesnr & "'"
                            sSQL = sSQL & " and linr = " & lLinr
                            Set rsArtlief = gdBase.OpenRecordset(sSQL)
                            If Not rsArtlief.EOF Then
                            
                                If Not IsNull(rsArtlief!artnr) Then
                                    sArtnr = rsArtlief!artnr
                                Else
                                    sArtnr = ""
                                End If
                                
                                If sArtnr <> "" Then
                        
                                    sSQL = "select * from artikel where artnr = " & sArtnr
                                    Set rsEAN = gdBase.OpenRecordset(sSQL)
                                    
                                    rsRs2.AddNew
                                    If Not rsEAN.EOF Then
                                        sSTATUS = "bekannt"
                                
                                        rsRs2!artnr = rsEAN!artnr
                                        rsRs2!BEZEICH = rsEAN!BEZEICH
                                        rsRs2!EAN = rsEAN!EAN
                                        rsRs2!Reihenf = i
                                        rsRs2!Status = sSTATUS
                                        rsRs2!AGN = rsEAN!AGN
                                        rsRs2!linr = rsArtlief!linr
                                        rsRs2!LPZ = rsEAN!LPZ
                                        rsRs2!LIBESNR = rsArtlief!LIBESNR
                                        rsRs2!RKZ = rsEAN!RKZ
                                        rsRs2!NOTIZEN = rsEAN!NOTIZEN
                                        rsRs2!BESTAND = rsEAN!BESTAND
                                        rsRs2!kvkpr = rsEAN!KVKPR1
                                        rsRs2!LEK = rsArtlief!lekpr
                                        rsRs2!Menge = CLng(sVPEMENGE) * rsArtlief!MINMEN
                                        rsRs2!ETIMERK = rsEAN!ETIMERK
                                        rsRs2!nsn = rsEAN!SPANNE
                                     Else
                                        sSTATUS = "unbekannt"
                                        rsRs2!LIBESNR = rsrs!LIBESNR
                                        rsRs2!Menge = sVPEMENGE
                                        rsRs2!Reihenf = i
                                        rsRs2!Status = sSTATUS
                                    End If
                                    
                                    rsRs2.Update
                                    
                                    rsEAN.Close
                                Else
                                
                                    rsRs2.AddNew
                                    sSTATUS = "unbekannt"
                                        
                                    rsRs2!LIBESNR = rsrs!LIBESNR
                                    rsRs2!Menge = sVPEMENGE
                                    rsRs2!Reihenf = i
                                    rsRs2!Status = sSTATUS
                                    
                                    rsRs2.Update
                                End If
                            Else
                                
                                rsRs2.AddNew
                                sSTATUS = "unbekannt"
                                    
                                rsRs2!LIBESNR = rsrs!LIBESNR
                                rsRs2!Menge = sVPEMENGE
                                rsRs2!Reihenf = i
                                rsRs2!Status = sSTATUS
                                
                                rsRs2.Update
                            
                            End If
                            rsArtlief.Close: Set rsArtlief = Nothing
                        End If



                    End If
                End If
            Loop While lLenfil >= lPos
        End If
        
        Close iFileNr

        rsRs2.Close: Set rsRs2 = Nothing
        
        Screen.MousePointer = 0
    
        Excelimport = True
    End If
    
    
    
    
    
    
    
    
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3125 Or err.Number = 3011 Then
        anzeige "rot", "Die Excelliste hat nicht das erwartete Format", Label4
        Exit Function
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Excelimport"
        Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Function

Private Sub DESADV_Uebereinstimmung(sANummer As String, sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsMDE       As Recordset
    Dim rsMDE1      As Recordset
    Dim rsEAN       As Recordset
'    Dim rsArtlief   As Recordset
    Dim sArtnr      As String
    Dim sSTATUS     As String
    Dim i           As Integer
    
    
    i = 0
    loeschNEW "WAEINGEM", gdBase
    CreateTableT2 "WAEINGEM", gdBase
    

    
    Set rsMDE = gdBase.OpenRecordset("WAEINGEM", dbOpenTable)
    
    sSQL = "Select * from DESADV where auftragsnr = " & sANummer & " order by libesnr "
    Set rsMDE1 = gdBase.OpenRecordset(sSQL)
    
    If Not rsMDE1.EOF Then
        rsMDE1.MoveFirst
        Do While Not rsMDE1.EOF
            i = i + 1
            If Not IsNull(rsMDE1!artnr) Then
                sArtnr = Trim(rsMDE1!artnr)
            Else
                sArtnr = ""
            End If
            
            If sArtnr <> "" Then
                
                sSQL = "Select "
                sSQL = sSQL & "artikel.artnr"
                sSQL = sSQL & ", artikel.BEZEICH"
                sSQL = sSQL & ", artikel.EAN"
                sSQL = sSQL & ", artikel.AGN"
                sSQL = sSQL & ", Artlief.linr "
                sSQL = sSQL & ", artikel.LPZ"
                sSQL = sSQL & ", Artlief.LIBESNR"
                sSQL = sSQL & ", Artlief.RKZ"
                sSQL = sSQL & ", artikel.NOTIZEN"
                sSQL = sSQL & ", artikel.BESTAND"
                sSQL = sSQL & ", artikel.KVKPR1"
                sSQL = sSQL & ", artikel.ETIMERK"
                sSQL = sSQL & ", artikel.Spanne"
                sSQL = sSQL & ", Artlief.lekpr"
                
                
                
                
                
                
                sSQL = sSQL & " from artikel inner join ARTLIEF on artikel.artnr = ARTLIEF.ARTNR"
                sSQL = sSQL & " where ARTLIEF.LINR = " & Val(sLinr) & " and artikel.artnr = " & sArtnr & " "
                Set rsEAN = gdBase.OpenRecordset(sSQL)
                
                rsMDE.AddNew
                If Not rsEAN.EOF Then
                    sSTATUS = "bekannt"
            
                    rsMDE!artnr = rsEAN!artnr
                    rsMDE!BEZEICH = rsEAN!BEZEICH
                    rsMDE!EAN = rsEAN!EAN
                    rsMDE!Menge = rsMDE1!Menge
                    rsMDE!Reihenf = i
                    rsMDE!Status = sSTATUS
                    rsMDE!AGN = rsEAN!AGN
                    rsMDE!linr = rsEAN!linr
                    rsMDE!LPZ = rsEAN!LPZ
                    rsMDE!LIBESNR = rsEAN!LIBESNR
                    rsMDE!RKZ = rsEAN!RKZ
                    rsMDE!NOTIZEN = rsEAN!NOTIZEN
                    rsMDE!BESTAND = rsEAN!BESTAND
                    rsMDE!kvkpr = rsEAN!KVKPR1
                    
                    rsMDE!LEK = rsEAN!lekpr
                    
'                    sSQL = "Select * from artlief where artnr = " & rsEAN!artnr
'                    sSQL = sSQL & " and linr = " & rsEAN!linr
'                    Set rsArtlief = gdBase.OpenRecordset(sSQL)
'                    If Not rsArtlief.EOF Then
'                        rsMDE!LEK = rsArtlief!lekpr
'                    Else
'
'                    End If
'                    rsArtlief.Close: Set rsArtlief = Nothing
                    
                    rsMDE!ETIMERK = rsEAN!ETIMERK
                    rsMDE!nsn = rsEAN!SPANNE
                 Else
                    sSTATUS = "unbekannt"
                    
                    rsMDE!LIBESNR = rsMDE1!LIBESNR
                    rsMDE!Menge = rsMDE1!Menge
                    rsMDE!Reihenf = i
                    rsMDE!Status = sSTATUS
                End If
                
                rsEAN.Close
            Else
            
                rsMDE.AddNew
                sSTATUS = "unbekannt"
                    
                rsMDE!LIBESNR = rsMDE1!LIBESNR
                rsMDE!Menge = rsMDE1!Menge
                rsMDE!Reihenf = i
                rsMDE!Status = sSTATUS
            End If
            
            rsMDE.Update
            rsMDE1.MoveNext
        Loop
    End If
    
    rsMDE1.Close: Set rsMDE1 = Nothing
    rsMDE.Close: Set rsMDE = Nothing
    
    
    
    
    
    
    
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DESADV_Uebereinstimmung"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Sub DESADV_anwenden_Einzel(sANummer As String, sLinr As String)
On Error GoTo LOKAL_ERROR

    Frame4.Visible = True
    Command8.Visible = False
    Command9.Visible = False
    Command10.Visible = False
    Command13.Visible = False
    Command12.Visible = False
    Label6(0).Visible = False
    Label6(1).Visible = False
    Label6(2).Visible = False
    Label6(3).Visible = False
    Command8.Enabled = True
    Command9.Enabled = True
    
    Command8.BackColorFrom = glButtonHintergrund_from
    Command8.BackColorTo = glButtonHintergrund_to

    Label4(0).Caption = "Daten aus der Datei werden ausgelesen..."
    Label4(0).Refresh

    DESADV_Uebereinstimmung sANummer, sLinr
    anzeigenalleDS

    Screen.MousePointer = 0
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DESADV_anwenden_Einzel"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DESADV_anwenden()
On Error GoTo LOKAL_ERROR

    Dim sAuftragsnr     As String

    If List2.ListIndex < 0 Then
        MsgBox "Bitte eine Datei auswählen!", vbInformation, "Winkiss Hinweis:"
        List2.SetFocus
    Else
        sAuftragsnr = List2.list(List2.ListIndex)
        sAuftragsnr = Trim(Mid$(sAuftragsnr, 1, InStr(1, sAuftragsnr, " ")))
        
        DESADV_anwenden_Einzel sAuftragsnr, Label19.Caption
        Frame2.Visible = False

    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DESADV_anwenden"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub del_DESADV()
On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim iRet            As Integer
    Dim sAuftragsnr     As String

    If List2.ListIndex < 0 Then
        iRet = MsgBox("Möchten Sie alle Dateien löschen?", vbYesNo + vbQuestion + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            For i = 0 To List2.ListCount - 1
                sAuftragsnr = List2.list(i)
                sAuftragsnr = Trim(Mid$(sAuftragsnr, 1, InStr(1, sAuftragsnr, " ")))
                
                del_DESADV_Einzel sAuftragsnr
            Next i
        End If
    Else
        sAuftragsnr = List2.list(List2.ListIndex)
        sAuftragsnr = Trim(Mid$(sAuftragsnr, 1, InStr(1, sAuftragsnr, " ")))
        
        del_DESADV_Einzel sAuftragsnr
    End If

    fuelle_Frame2_mit_DESADV
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "del_DESADV"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "PRINT_WE", gdBase
    CreateTableT2 "PRINT_WE", gdBase
    
    sSQL = "Insert into PRINT_WE Select * from WAEINGEM"
    
    iRet = MsgBox("Möchten Sie nur die unbekannten Artikel einsehen?", vbInformation + vbYesNo, "Winkiss Frage:")
    If iRet = vbYes Then
        'unbekannte
        sSQL = sSQL & " where status = 'unbekannt'"
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    reportbildschirm "dWKL15", "aWKL15a"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
        
End Sub
Private Sub Command10_Click()
    On Error GoTo LOKAL_ERROR
    
    Command5_Click
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Label4(0).Caption = "Hier können Sie ein Übereinstimmungsprotokoll einsehen."
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Label4(0).Caption = "Hier können Sie unbekannte EAN - Codes zu bestehenden Artikel in der Datenbank zuordnen."
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command13_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Label4(0).Caption = "Hier werden alle erkannten Artikel dem Bestand zugebucht."
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo LOKAL_ERROR
    
    Label4(0).Caption = "Hier können Sie das Protokoll des Wareneingangs ansehen/drucken."
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command10_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command9_Click()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMDE   As Recordset
    Dim sEAN    As String
    Dim sMenge  As String
    Dim sSQL    As String
    Dim iRet    As Integer
    Dim sLEK    As String
    Dim sKVKPR  As String
    
    Screen.MousePointer = 11
    
    If Trim$(Combo1.Text) = "" Then 'Frage nach LS Nummer
        iRet = MsgBox("Möchten Sie eine Lieferschein-Nr. eingeben?", vbCritical + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
        If iRet = vbYes Then
            Combo1.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        Else
            Combo1.Text = "keine"
        End If
    End If

    Command8.Enabled = False
    Command9.Enabled = False
    
    sSQL = "Delete from WAEINGEM where status = 'unbekannt' "
    gdBase.Execute sSQL, dbFailOnError
    
    MDEduplikatsverarbeitung
    
    Dim iCount As Integer
    
    Set rsMDE = gdBase.OpenRecordset("MDEINH", dbOpenTable)
    If Not rsMDE.EOF Then
        rsMDE.MoveLast
        iCount = rsMDE.RecordCount
        rsMDE.MoveFirst
        Do While Not rsMDE.EOF
        
        
            Label4(0).Caption = "noch " & iCount
            iCount = iCount - 1
            
            If Not IsNull(rsMDE!EAN) Then
                sEAN = rsMDE!EAN
            Else
                sEAN = ""
            End If

            If Not IsNull(rsMDE!MENG) Then
                sMenge = rsMDE!MENG
            Else
                sMenge = ""
            End If
            Text1(0).Text = Trim(sEAN)
            
            If Not IsNull(rsMDE!LEK) Then
                 sLEK = rsMDE!LEK
            Else
                sLEK = ""
            End If
            
            If Not IsNull(rsMDE!kvkpr) Then
                sKVKPR = rsMDE!kvkpr
            Else
                sKVKPR = ""
            End If
            Text1(0).Text = Trim(sEAN)
            
            fromMde = True
            Command1_Click 'SuchenButton
            Me.Refresh
            If Label2(0).Caption <> "Artikel nicht gefunden!" Then
                
                Text1(1).Text = sMenge
                Text1(3).Text = Format$(sLEK, "######0.00")
                Text1(2).Text = Format$(sKVKPR, "######0.00")
                Command2_Click 15 'SpeichernButton
            End If

            rsMDE.MoveNext
        Loop
    End If
    rsMDE.Close: Set rsMDE = Nothing
    
    Label4(0).Caption = "Daten aus dem MDE - Gerät / DESADV (Lieferavis) sind erfolgreich verarbeitet!"
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command9_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim cSQL As String
    Dim i As Integer
    
    bfoundauto = False
    fromMde = False
    bscanner = False
    gbLibesnrSeek = False
   
    PositionierenWKL15
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    Frame2.BackColor = glH2
    Label16.BackColor = glH2
    Label17.BackColor = glH2
    Label18.BackColor = glH2
    Label19.BackColor = glH2
    
    Command0(0).BackColorFrom = vbWhite
    Command0(0).BackColorTo = vbWhite
    
    If gbscanmodi Then
        Check9.Value = vbChecked
    Else
        Check9.Value = vbUnchecked
    End If
    
    For i = 0 To 6
        Label15(i).ForeColor = vbBlack
        Label15(i).Caption = ""
    Next i
    
    If Not NewTableSuchenDBKombi("ETIDRULS", gdBase) Then
        CreateTableT2 "ETIDRULS", gdBase
    End If
    
    LeseLieferschein "ETIDRULS", Combo1

    Check8.Value = vbUnchecked

    LeereDialogWKL15
    gF2Prompt.lLastPos = -1
    
    glSelect = 0
    lSelect = 0
    
    If gsWeEinzFo = "EAN" Then
        Text1(0).TabIndex = 0
        Combo1.Text = "keine"
    Else
        Combo1.TabIndex = 0
    End If
    
    Text1(1).Text = gsWeEinzMe
    
    List1.Clear
    List1.AddItem " Artnr Artikelbezeichnung                 Menge       REK      LEK      SEK      KVK"
    
    combofuell Combo2
    
    If Modul6.FindFile(gcDBPfad, "aWKL30ys.rpt") Then
        Check2.Visible = True
        Check2.Value = vbUnchecked
    ElseIf Modul6.FindFile(App.Path, "aWKL30xs.rpt") Then
        Check2.Visible = True
        Check2.Value = vbUnchecked
    
    Else
        Check2.Visible = False
        Check2.Value = vbUnchecked
    End If
    
    If Modul6.FindFile(gcDBPfad, "aLFNR.rpt") Then
        Check3.Visible = True
        Check3.Value = vbUnchecked
        
    ElseIf Modul6.FindFile(App.Path, "aWKL30zs.rpt") Then
        Check3.Visible = True
        Check3.Value = vbUnchecked
    Else
        Check3.Visible = False
        Check3.Value = vbUnchecked
    End If
    
    If gbETIKVKAE Then
        Check65.Value = vbChecked
    Else
        Check65.Value = vbUnchecked
    End If
    
    If gbNONEGZU Then 'keine negativen
        Command2(12).Enabled = False
    Else
        Command2(12).Enabled = True
    End If
    
    Label2(9).Caption = ""
    
    If NewTableSuchenDBKombi("E15B", gdApp) Then
        If SpalteInTabellegefundenNEW("E15B", "BO2", gdApp) = False Then
            SpalteAnfuegenNEW "E15B", "BO2", "Bit", gdApp
        
        End If
        voreinstellungladenE15B
    
    End If
    
    bLieferavis = False
    
    Screen.MousePointer = 0
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub voreinstellungladenE15B()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Set rs = gdApp.OpenRecordset("E15B")
    If Not rs.EOF Then
    
        If rs!bo1 = True Then
            Check2.Value = vbUnchecked
        Else
            Check2.Value = vbChecked
        End If
        
        If rs!bo2 = True Then
            Option1(0).Value = True
        Else
            Option1(1).Value = True
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE15B"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE15B()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim bo1     As Integer
    Dim bo2     As Integer
    
    loeschNEW "E15B", gdApp
    CreateTable "E15B", gdApp
    
    If Check2.Value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If
    
    If Option1(1).Value = True Then
        bo2 = 0
    Else
        bo2 = -1
    End If
    
    sSQL = "Insert into E15B ( bo1,bo2) "
    sSQL = sSQL & " values (" & bo1 & "," & bo2 & ")"
    gdApp.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE15B"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub combofuell(cbo As ComboBox)
    On Error GoTo LOKAL_ERROR

    cbo.Clear
    cbo.Visible = True
    cbo.AddItem "1"
    cbo.AddItem "2"
    cbo.AddItem "3"
    cbo.AddItem "4"
    cbo.AddItem "5"
    cbo.AddItem "6"
    cbo.AddItem "7"
    cbo.AddItem "8"
    cbo.AddItem "9"
    cbo.AddItem "10"
    cbo.AddItem "15"
    cbo.AddItem "20"
    cbo.AddItem "25"
    cbo.AddItem "30"
    cbo.AddItem "40"
    cbo.AddItem "50"
    
    cbo.Text = ""
    cbo.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "combofuell"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label2(7).ForeColor = glS1
    Label1(18).ForeColor = glS1
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo LOKAL_ERROR
    
    loeschNEW "WAEINGEM", gdBase
    loeschNEW "PRINT_WE", gdBase
    
    voreinstellungspeichernE15B
    LogtoEnd Me
    Unload Me
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(19).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame3_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Image2_Click()
    On Error GoTo LOKAL_ERROR
    Dim lRet As Long
    
    MSHFLEX1.Clear
    MSHFLEX1.Visible = False
    Frame4.Visible = False
    Frame7.Visible = False
    
    NachLinrEingabeWeiter
    
    lblUeberschrift.ForeColor = glU1
    lblUeberschrift.Caption = "Wareneingang aus Einzellieferung"
    lblUeberschrift.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Image2_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Function fnPruefeLINRWKL15(sLinr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cLinr As String
    Dim cSQL As String
    Dim rsrs As Recordset
    
    fnPruefeLINRWKL15 = True
    
    cLinr = Trim$(Str$(Val(sLinr)))
    
    If cLinr = "0" Then
        fnPruefeLINRWKL15 = False
        Exit Function
    End If
    
    If cLinr = "" Then
        fnPruefeLINRWKL15 = False
        Exit Function
    End If
    
    cSQL = "Select * from LISRT where LINR = " & cLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        fnPruefeLINRWKL15 = False
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeLINRWKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Sub NachLinrEingabeWeiter()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim sLinr   As String
    
    Frame4.Visible = True
    Command8.Visible = False
    Command9.Visible = False
    Command10.Visible = False
    Command13.Visible = False
    Command12.Visible = False
    Label6(0).Visible = False
    Label6(1).Visible = False
    Label6(2).Visible = False
    Label6(3).Visible = False
    Command8.Enabled = True
    Command9.Enabled = True
    
    Command8.BackColorFrom = glButtonHintergrund_from
    Command8.BackColorTo = glButtonHintergrund_to

    Label4(0).Caption = "Daten aus dem MDE - Gerät werden ausgelesen..."
    Label4(0).Refresh

    If MDEeinlesenOhneLinr(Label14, txtStatus, picprogress, frmWKL15) = False Then
        Label4(0).Caption = "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden."
    Else
        MDEuebereinstimmung
        anzeigenalleDS
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NachLinrEingabeWeiter"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "MENGE"
                SpaltennummerMENGE = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub anzeigenalleDS()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsUn    As Recordset
    
    Tabcheck "WAEINGEM"
    
    FormatMShFlex1WKL15     'Tabelle erstellen
    FuellenMShFlex1WKL15    'Tabelle füllen
    
    ermittlespalten
    
    Tabellenbreiteanpassen MSHFLEX1, 1.25 * gdTabfak
    
    sSQL = " Select * from WAEINGEM where Status = 'unbekannt' "
    Set rsUn = gdBase.OpenRecordset(sSQL)
    If Not rsUn.EOF Then
        Command13.Enabled = True
        Command8.BackColorFrom = vbRed
        Command8.BackColorTo = vbRed
    Else
        Command13.Enabled = False
        
    End If
    rsUn.Close
    Command8.Visible = True
    Command9.Visible = True
    Command10.Visible = True
    Command13.Visible = True
    Command12.Visible = True
    Label6(0).Visible = True
    Label6(1).Visible = True
    Label6(2).Visible = True
    Label6(3).Visible = True
    Label4(0).Caption = "Wählen sie den nächsten Schritt!"
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "anzeigenalleDS"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
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
Private Sub FuellenMShFlex1WKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim i           As Integer
    Dim j           As Integer

    Set rsrs = gdBase.OpenRecordset("WAEINGEM", dbOpenTable)

    MSHFLEX1.Redraw = False
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            
            
            lrow = lrow + 1

            MSHFLEX1.Rows = lrow + 1
            MSHFLEX1.Col = 0
 
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Row = 0
                MSHFLEX1.Col = i

                If sSpaltenname(i) = MSHFLEX1.Text Then

                    Select Case sSpaltenname(i)
                        Case Is = "Listen - EK", "Kassen - Preis"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = Format$(sWert, "####0.00")
                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            MSHFLEX1.Row = lrow
                            MSHFLEX1.Text = sWert
                    End Select


                    If TextWidth(MSHFLEX1.TextMatrix(lrow, i)) > aBreite(i) Then
'                        aBreite(i) = Len(MSHFLEX1.TextMatrix(0, i)) * 80
                        aBreite(i) = TextWidth(MSHFLEX1.TextMatrix(lrow, i))
                    End If

                End If
            Next i
            
            If Not IsNull(rsrs!Status) Then
                sWert = rsrs!Status
            End If
            
            For j = 1 To byAnzahlSpalten - 1
                MSHFLEX1.Col = j
                If sWert <> "bekannt" Then
                    MSHFLEX1.CellBackColor = vbRed
                Else
                    MSHFLEX1.CellBackColor = vbWhite
                End If
            Next j
            rsrs.MoveNext
        Loop
    End If

    For i = 0 To byAnzahlSpalten - 1
        MSHFLEX1.Col = i
        MSHFLEX1.ColWidth(i) = aBreite(i) * 1.5
    Next i


    rsrs.Close: Set rsrs = Nothing
    If byAnzahlSpalten < 2 Then

    Else
        MSHFLEX1.FixedCols = 1
    End If

    MSHFLEX1.RowHeight(1) = 0
    lrow = lrow - 1


    MSHFLEX1.Redraw = True
    MSHFLEX1.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex1WKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
    
End Sub
Private Sub FuellenMShFlex2WKL15a()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim i           As Integer
    Dim sSQL        As String
    
    sSQL = " Select * from WAEINGEM where Status = 'unbekannt' "
    Set rsrs = gdBase.OpenRecordset(sSQL)

    sSQL = "Delete from Artikel where bezeich is null or bezeich = ' ' or bezeich ='' "
    gdBase.Execute sSQL, dbFailOnError

    MSHFlex2.Redraw = False
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1

            MSHFlex2.Rows = lrow + 1
            MSHFlex2.Row = lrow
            
            MSHFlex2.Col = 0
            MSHFlex2.Text = IIf(IsNull(rsrs!Reihenf), "", rsrs!Reihenf)
            
            MSHFlex2.Col = 1
            MSHFlex2.Text = IIf(IsNull(rsrs!Menge), "", rsrs!Menge)
            
            MSHFlex2.Col = 2
            MSHFlex2.Text = IIf(IsNull(rsrs!EAN), "", rsrs!EAN)
            
            MSHFlex2.Col = 5
            MSHFlex2.Text = HoleFreieArtikelNrWKL10
            
            sSQL = "Insert into artikel (artnr) values ('" & MSHFlex2.Text & "') "
            gdBase.Execute sSQL, dbFailOnError
            
            rsrs.MoveNext
        Loop
    End If

    For i = 0 To byAnzahlSpalten - 1
        If TextWidth(MSHFlex2.TextMatrix(lrow, i)) > aBreite(i) Then
            aBreite(i) = TextWidth(MSHFlex2.TextMatrix(lrow, i))
        End If
            
        MSHFlex2.Col = i
        MSHFlex2.ColWidth(i) = aBreite(i) * 1.5
    Next i

    rsrs.Close: Set rsrs = Nothing
    If byAnzahlSpalten < 2 Then

    Else
        MSHFlex2.FixedCols = 1
    End If

    MSHFlex2.RowHeight(1) = 0
    lrow = lrow - 1


    MSHFlex2.Redraw = True
    MSHFlex2.Visible = True
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FuellenMShFlex2WKL15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub FormatMShFlex2WKL15a()
On Error GoTo LOKAL_ERROR

    Dim i As Byte
    Dim j As Byte
    
    MSHFLEX1.Visible = False
    byAnzahlSpalten = 14
    
    ReDim sSpaltenname(byAnzahlSpalten)
    ReDim sSpaltenbez(byAnzahlSpalten)
    ReDim aBreite(byAnzahlSpalten)
    
    sSpaltenname(0) = "Reihenfolge"
    sSpaltenbez(0) = "Reihenf"
    
    sSpaltenname(1) = "Menge"
    sSpaltenbez(1) = "Menge"
    
    sSpaltenname(2) = "EAN"
    sSpaltenbez(2) = "EAN"
    
    sSpaltenname(3) = "2.EAN"
    sSpaltenbez(3) = "EAN2"
    
    sSpaltenname(4) = "3.EAN"
    sSpaltenbez(4) = "EAN3"
    
    sSpaltenname(5) = "Artnr"
    sSpaltenbez(5) = "Artnr"
    
    sSpaltenname(6) = "Artikelbezeichnung"
    sSpaltenbez(6) = "BEZEICH"
    
    sSpaltenname(7) = "LieferantenNr"
    sSpaltenbez(7) = "Linr"
    
    sSpaltenname(8) = "Linie"
    sSpaltenbez(8) = "Lpz"
    
    sSpaltenname(9) = "Lieferantenbestnr."
    sSpaltenbez(9) = "Libesnr"
    
    sSpaltenname(10) = "EK - Preis"
    sSpaltenbez(10) = "Ekpr"

    sSpaltenname(11) = "VK - Preis"
    sSpaltenbez(11) = "KVKPR1"
    
    sSpaltenname(12) = "AGN"
    sSpaltenbez(12) = "AGN"
    
    sSpaltenname(13) = "MwSt"
    sSpaltenbez(13) = "MWST"
    
    
    

    With MSHFlex2
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With
        
Exit Sub
LOKAL_ERROR:
Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatMShFlex2WKL15a"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FormatMShFlex1WKL15()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsRL As Recordset
    
    Dim i As Byte
    Dim j As Byte
    
    MSHFlex2.Visible = False
    
    sSQL = "Select * from TABLay" & srechnertab & " where ANZEIGE = 'J' and Tabname = 'WAEINGEM' order by Reihenf"
    Set rsRL = gdBase.OpenRecordset(sSQL)
    
    If Not rsRL.EOF Then
        byAnzahlSpalten = rsRL.RecordCount
        ReDim sSpaltenname(byAnzahlSpalten)
        ReDim sSpaltenbez(byAnzahlSpalten)
        ReDim aBreite(byAnzahlSpalten)
        rsRL.MoveFirst
        i = 0
        Do While Not rsRL.EOF
            sSpaltenname(i) = rsRL!Spaltenna
            sSpaltenbez(i) = rsRL!Spaltenbez
            i = i + 1
            rsRL.MoveNext
        Loop
    End If
    rsRL.Close
    
    With MSHFLEX1
        .Redraw = False
        .Visible = False
        .Clear
        .Rows = 25
        .Cols = byAnzahlSpalten
        .FixedCols = 0
        .FixedRows = 1
        .Row = 0
        
        For j = 0 To byAnzahlSpalten - 1
            .Col = j
            .Text = sSpaltenname(j)
            aBreite(j) = TextWidth(.TextMatrix(0, j)) ' * 1.8
        Next j
    End With

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatMShFlex1WKL15"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_DblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

    If Index = 19 Then
        frmWKL192.Show 1
        
    ElseIf Index = 18 Then
    
        
        'check doch mal ob es Budni ist
        'wenn ja dann check mal ob ein Lieferavis vorliegt
        check_Auf_Budni_Lieferavis
        
        Frame2.Visible = True
        
        fuelle_Frame2_mit_DESADV
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub check_Auf_Budni_Lieferavis()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim sBudniKundnr    As String
    Dim sCheckLinr      As String
    Dim rsLi            As DAO.Recordset
    
    sBudniKundnr = ""
    sCheckLinr = ""
    
    sSQL = "select KUNDNR,linr from LISRT where FORMAT = 'EDIBUDNI' "
    Set rsLi = gdBase.OpenRecordset(sSQL)
    If Not rsLi.EOF Then
        sBudniKundnr = Trim(rsLi!Kundnr)
        sCheckLinr = Trim(rsLi!linr)
        
        Label19.Caption = sCheckLinr
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
     Fehler.gsFunktion = "check_Auf_Budni_Lieferavis"
     Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
     
     Fehlermeldung1

End Sub

Private Sub fuelle_Frame2_mit_DESADV()
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsrs    As DAO.Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    
    cSQL = "Select distinct(auftragsnr) as aufnr from DESADV order by auftragsnr desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    List2.Clear
    
    If Not rsrs.EOF Then
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!aufnr) Then
                cFeld = rsrs!aufnr
            End If
    
            cLBSatz = cFeld & Space(7 - Len(cFeld))
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
'    setzefokus Label3(10).Caption

    List2.Refresh
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelle_Frame2_mit_DESADV"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function pfadseekExcel_WE() As String
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String
    
    pfadseekExcel_WE = ""

    sTitle = "Speichern des Pfades"
    
    sFilter = "Excel - Dateien (*.xls)|*.xls"
    
    sOldpfad = gcDBPfad & "\IN"
    pfadseekExcel_WE = pfadaendernKomplett(sTitle, sFilter, sOldpfad)
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcel_WE"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If Index = 19 Then
        Label1(19).ForeColor = glLink
    ElseIf Index = 18 Then
        Label1(18).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label2_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim iRet As Integer

    If Index = 7 Then
        If Label2(7).Caption = "Stück" Then
            iRet = MsgBox("Möchten Sie wirklich auf Gewicht(Kg) umstellen", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbYes Then
                Label2(7).Caption = "Kg"
                Text1(1).MaxLength = 8
                
                List1.Clear
                List1.AddItem " Artnr Artikelbezeichnung                       Kilo"
                
                If Not NewTableSuchenDBKombi("KILOART", gdBase) Then
                    CreateTableT2 "KILOART", gdBase
                End If
                
            End If
        Else
            Label2(7).Caption = "Stück"
            Text1(1).MaxLength = 5
            
            List1.Clear
            List1.AddItem " Artnr Artikelbezeichnung                  Menge REK      LEK      SEK      KVK"
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label2_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If Index = 7 Then
        Label2(7).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label2_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label4_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 32 Then
    Label4(Index).Caption = "alle Farben"
    Label4(Index).Tag = ""
    Label4(Index).BackColor = Label11(4).BackColor
    Label4(Index).ForeColor = Label11(4).ForeColor
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_dblClick"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_Click()
    On Error GoTo LOKAL_ERROR
    
    glSelect = MSHFLEX1.Row
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX1_DblClick()
    On Error GoTo LOKAL_ERROR
    
    glSelect = MSHFLEX1.Row
    Command13_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub posimerk()
    On Error GoTo LOKAL_ERROR
    
    If MSHFLEX1.Row = 1 Then
        label0(0).Caption = "1"
        Exit Sub
    End If
    
    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSHFLEX1.Row
    lcol = MSHFLEX1.Col
    
    If lrow < 1 Then
        lrow = 1
    End If
    If lrow = MSHFLEX1.Rows Then
        lrow = lrow - 1
    End If
   
    MSHFLEX1.Row = lrow
    MSHFLEX1.Col = lcol
    
    label0(0).Caption = Trim$(Str$(lrow))
    label0(1).Caption = Trim$(Str$(lcol))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "posimerk"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long
    
    lrow = MSHFLEX1.Row
    lcol = MSHFLEX1.Col
    
    If KeyCode <> vbKeyDown And KeyCode <> vbKeyUp And KeyCode <> vbKeyRight And KeyCode <> vbKeyLeft Then  'And KeyCode <> vbKeyReturn
    
        Select Case lcol
            Case Is = SpaltennummerMENGE
                If iKeypress = 0 And KeyCode <> vbKeyBack And KeyCode <> vbKeyF2 And KeyCode <> vbKeyReturn Then
                    MSHFLEX1.Row = lrow
                    MSHFLEX1.Col = lcol
                    MSHFLEX1.Text = ""
                ElseIf iKeypress > 0 And KeyCode = 46 Then
                    MSHFLEX1.Row = lrow
                    MSHFLEX1.Col = lcol
                    MSHFLEX1.Text = ""
                End If
                iKeypress = iKeypress + 1
        End Select
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL                As String
    Dim cZeichen            As String
    Dim cValid              As String
    Dim sEAN                As String
    Dim lcol                As Long
    Dim lrow                As Long
    Dim bLEKVeraenderung    As Boolean
    Dim bKVKVeraenderung    As Boolean
    Dim bBestandsVeraenderung    As Boolean
    Dim i                   As Integer
    
    posimerk

    cZeichen = Chr$(KeyAscii)

    bLEKVeraenderung = True
    bKVKVeraenderung = False
    bBestandsVeraenderung = False
    
    MSHFLEX1.Redraw = False

    MSHFLEX1.Row = 0
    Select Case MSHFLEX1.Text
        Case Is = "Listen - EK"
            bLEKVeraenderung = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Col = i
                If MSHFLEX1.Text = "EAN" Then
                    lrow = Val(label0(0).Caption)
                    MSHFLEX1.Row = lrow
                    sEAN = MSHFLEX1.Text
                    Exit For
                End If
            Next i
                
        Case Is = "Kassen - Preis"
            bLEKVeraenderung = False
            bKVKVeraenderung = True
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Col = i
                If MSHFLEX1.Text = "EAN" Then
                    lrow = Val(label0(0).Caption)
                    MSHFLEX1.Row = lrow
                    sEAN = MSHFLEX1.Text
                    Exit For
                End If
            Next i
        
        Case Is = "Menge"
            bLEKVeraenderung = False
            bBestandsVeraenderung = True
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            For i = 0 To byAnzahlSpalten - 1
                MSHFLEX1.Col = i
                If MSHFLEX1.Text = "EAN" Then
                    lrow = Val(label0(0).Caption)
                    MSHFLEX1.Row = lrow
                    sEAN = MSHFLEX1.Text
                    Exit For
                End If
            Next i
            

        
        Case Else
            bLEKVeraenderung = False
            KeyAscii = 0
    End Select

    lcol = Val(label0(1).Caption)
    lrow = Val(label0(0).Caption)
    MSHFLEX1.Row = lrow
    MSHFLEX1.Col = lcol
    
    If KeyAscii <> 0 Then
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        cValid = MSHFLEX1.Text
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
            MSHFLEX1.Text = UCase(cValid)
        End If

    End If
    
    
    If bLEKVeraenderung Then
        
        MSHFLEX1.Row = 0
        For i = 0 To byAnzahlSpalten - 1
            MSHFLEX1.Col = i
            If MSHFLEX1.Text = "Artnr" Then
                lrow = Val(label0(0).Caption)
                MSHFLEX1.Row = lrow
                Label2(2).Caption = MSHFLEX1.Text
                Exit For
            End If
        Next i
        
        MSHFLEX1.Row = 0
        For i = 0 To byAnzahlSpalten - 1
            MSHFLEX1.Col = i
            If MSHFLEX1.Text = "Menge" Then
                lrow = Val(label0(0).Caption)
                MSHFLEX1.Row = lrow
                Text1(1).Text = MSHFLEX1.Text
                Exit For
            End If
        Next i
        
        Text1(3).Text = UCase(cValid)
        
        MSHFLEX1.Row = 0
        For i = 0 To byAnzahlSpalten - 1
            MSHFLEX1.Col = i
            If MSHFLEX1.Text = "Kassen - Preis" Then
                lrow = Val(label0(0).Caption)
                MSHFLEX1.Row = lrow
                MSHFLEX1.Text = Text1(2).Text
                Exit For
            End If
        Next i
        
        lcol = Val(label0(1).Caption)
        lrow = Val(label0(0).Caption)
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        
        If MSHFLEX1.Text <> "" Then
            sSQL = "Update WAEINGEM set LEK = '" & MSHFLEX1.Text & "'"
            sSQL = sSQL & " where ean = '" & sEAN & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    If bKVKVeraenderung Then
        lcol = Val(label0(1).Caption)
        lrow = Val(label0(0).Caption)
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        
        If MSHFLEX1.Text <> "" Then
            sSQL = "Update WAEINGEM set KVKPR = '" & MSHFLEX1.Text & "'"
            sSQL = sSQL & " where ean = '" & sEAN & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    If bBestandsVeraenderung Then
        lcol = Val(label0(1).Caption)
        lrow = Val(label0(0).Caption)
        MSHFLEX1.Row = lrow
        MSHFLEX1.Col = lcol
        
        If MSHFLEX1.Text <> "" Then
            sSQL = "Update WAEINGEM set Menge = '" & MSHFLEX1.Text & "'"
            sSQL = sSQL & " where ean = '" & sEAN & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
    End If
    
    
    
    
    MSHFLEX1.Redraw = True

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Select Case KeyCode
        Case Is = 46    'Del
            MSHFLEX1.Text = ""
    End Select

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub MSHFLEX1_LeaveCell()
On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub MSHFlex2_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR

    Dim lrow As Long
    Dim lcol As Long

    lrow = MSHFlex2.Row
    lcol = MSHFlex2.Col

    If iKeypress = 0 And KeyCode <> vbKeyBack Then
        MSHFlex2.Row = lrow
        MSHFlex2.Col = lcol
        MSHFlex2.Text = ""
    End If
    iKeypress = iKeypress + 1

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlex2_KeyDown"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub MSHFLEX2_KeyPress(KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcol     As Long
    Dim lrow     As Long
    Dim cZeichen As String
    Dim cFeld    As String

    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)

    cFeld = MSHFlex2.Text

    Select Case KeyAscii
        Case Is = 8
            If Len(cFeld) > 0 Then
                cFeld = Left$(cFeld, Len(cFeld) - 1)
            End If
        Case Else
            cFeld = cFeld & Chr$(KeyAscii)
    End Select

    MSHFlex2.TextMatrix(MSHFlex2.Row, MSHFlex2.Col) = cFeld
    MSHFlex2.Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlex2_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFLEX2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
     Select Case KeyCode
        Case Is = 46    'Del
            MSHFlex2.Text = ""
    End Select
    
   Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlex2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub MSHFLEX2_LeaveCell()
    On Error GoTo LOKAL_ERROR
    
    iKeypress = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFLEX2_LeaveCell"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MSHFlex3_Click()
    On Error GoTo LOKAL_ERROR
    
    lSelect = MSHFlex3.Row
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSHFlex3_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Option1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

'   voreinstellungspeichernE15B

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option1_Click"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub



Private Sub Text1_Change(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sNettospanne As String
    Dim dNettospanne As Double
    Dim i As Integer

    If Index = 3 Then
        If Trim(Label2(2).Caption) <> "" Then
            If Trim(Text1(4).Text) <> "" Then
                Text1(2).Text = ErmittleUeberNettospanneKVK(Trim(Label2(2).Caption), Trim(Text1(4).Text))
            End If
        End If
    End If
    
    If Index = 2 Or Index = 3 Then
    
        If Text1(2).Text <> "" And Text1(3).Text <> "" Then
            For i = 0 To 6
                Label15(i).Caption = ""
            Next i
        
            
            sNettospanne = NettospanneInProzent(Text1(2).Text, Text1(3).Text, Label2(10).Caption)
            
            dNettospanne = CDbl(sNettospanne)
            
            If dNettospanne >= 100 Then
                Label15(0).Caption = Fix(dNettospanne)
                Label15(0).ToolTipText = "Nettospanne: " & sNettospanne & " %"
            ElseIf dNettospanne > 79.99 Then
                Label15(1).Caption = Fix(dNettospanne)
                Label15(1).ToolTipText = "Nettospanne: " & sNettospanne & " %"
            ElseIf dNettospanne > 59.99 Then
                Label15(2).Caption = Fix(dNettospanne)
                Label15(2).ToolTipText = "Nettospanne: " & sNettospanne & " %"
            ElseIf dNettospanne > 39.99 Then
                Label15(3).Caption = Fix(dNettospanne)
                Label15(3).ToolTipText = "Nettospanne: " & sNettospanne & " %"
            ElseIf dNettospanne > 19.99 Then
                Label15(4).Caption = Fix(dNettospanne)
                Label15(4).ToolTipText = "Nettospanne: " & sNettospanne & " %"
            ElseIf dNettospanne > 0 Then
                Label15(5).Caption = Fix(dNettospanne)
                Label15(5).ToolTipText = "Nettospanne: " & sNettospanne & " %"
            ElseIf dNettospanne <= 0 Then
                Label15(6).Caption = Fix(dNettospanne)
                Label15(6).ToolTipText = "Nettospanne: " & sNettospanne & " %"
            End If
        End If
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ErmittleUeberNettospanneKVK(sArtikelnummer As String, sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsArt           As Recordset
    Dim rsArtlief       As Recordset
    Dim dEK             As Double
    Dim cMWST           As String
    Dim dNettospanne    As Double
    Dim se              As String
    
    Dim dEkPrAlt        As Double
    Dim dEkpr           As Double
    Dim dAnzahl         As Double
    Dim dWertNeu        As Double
    Dim dWertAlt        As Double
    Dim dAlt            As Double
    
    
    ErmittleUeberNettospanneKVK = Text1(2).Text
    
    sSQL = " Select artikel.* from artikel inner join artlief on artikel.artnr = artlief.artnr "
    sSQL = sSQL & " where artikel.artnr = " & sArtikelnummer
    sSQL = sSQL & " and artlief.linr = " & sLinr
    sSQL = sSQL & " and artikel.PREISSCHU <> 'J' "
    sSQL = sSQL & " and not artlief.spanne = 0 "
    
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
                
        If Not IsNull(rsArt!MWST) Then
            cMWST = rsArt!MWST
        Else
            cMWST = "V"
        End If
    
        sSQL = " Select * from artlief where artnr = " & sArtikelnummer
        sSQL = sSQL & " and Linr = " & sLinr

        Set rsArtlief = gdBase.OpenRecordset(sSQL)
        If Not rsArtlief.EOF Then
        
            If Not IsNull(rsArtlief!SPANNE) Then
                dNettospanne = rsArtlief!SPANNE
            Else
                dNettospanne = 0
            End If
        Else
        
        End If
        rsArtlief.Close: Set rsArtlief = Nothing
        
        Label12.Caption = Format$(dNettospanne, "#####0.00") & " %"
        Label12.Refresh
        
        If Not IsNull(rsArt!ekpr) Then 'Alter Schnitt EK
            dEkPrAlt = rsArt!ekpr
        Else
            dEkPrAlt = 0
        End If
        
        Label13.Caption = Format$(dEkPrAlt, "#####0.00") & Space(1) & gcWaehrung
        Label13.Refresh
            
        If Trim(Text1(3).Text) = "," Then Text1(3).Text = ""
        If gsSpanne = "LEK" Then
            If Text1(3).Text <> "" Then
                dEK = Text1(3).Text
            Else
                dEK = 0
            End If
            
            Label14.Caption = "Die Kalkulation bezieht auf den Listen - EK. "
            Label14.Refresh
            
        ElseIf gsSpanne = "SEK" Then
            
            If Text1(3).Text <> "" Then 'Listen EK aus Formular
                dEkpr = Text1(3).Text
            Else
                dEkpr = 0
            End If
            
            If Text1(1).Text <> "" Then
                dAnzahl = Val(Trim(Text1(1).Text))
            Else
                dAnzahl = 0
            End If
            
            If dAnzahl = 0 Then
                Label14.Caption = "Die Kalkulation bezieht auf den Schnitt - EK. "
                Label14.Caption = Label14.Caption & " Um den KVK - Preis zu verändern, müssen Sie 'Zu/Abgang' füllen! "
                Label14.Refresh
            Else
                Label14.Caption = "Die Kalkulation bezieht auf den Schnitt - EK. "
                Label14.Refresh
            End If
            If Not IsNull(rsArt!BESTAND) Then
                dAlt = rsArt!BESTAND
            Else
                dAlt = 0
            End If
            
            If dEkPrAlt <> 0 And dEkpr <> 0 Then
                 'Schnitt beginn
                
                dWertNeu = SchnittEKBerechnung(sArtikelnummer, CLng(sLinr), CLng(dAnzahl), dEkpr, CLng(dAlt))
                'Schnitt ende
            End If
            
            dEK = dWertNeu
            Label13.Caption = Format$(dWertNeu, "#####0.00") & Space(1) & gcWaehrung
            Label13.Refresh
        End If
        
        If Val(dEK) = 0 Then
            Text1(2).Text = ""
            lblUeberschrift.ForeColor = glU1
            lblUeberschrift.Caption = "Wareneingang aus Einzellieferung"
        Else
            ErmittleUeberNettospanneKVK = Runden(CDbl(fnVKneuNS(dEK, cMWST, dNettospanne)))
            lblUeberschrift.ForeColor = vbRed
            lblUeberschrift.Caption = "Automatische Kalkulation!"
        End If
    Else
        ErmittleUeberNettospanneKVK = Text1(2).Text
    End If
    rsArt.Close: Set rsArt = Nothing
    

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ErmittleUeberNettospanneKVK"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text1(Index).BackColor = glSelBack1
    Label3.Caption = Format$(Index, "##0")
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = glSelBack1
    Text3(Index).SelStart = 0
    Text3(Index).SelLength = Len(Text3(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo1_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Combo1.BackColor = glSelBack1
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "combo1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Combo2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Combo2.BackColor = glSelBack1
    Combo2.SelStart = 0
    Combo2.SelLength = Len(Combo2.Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo2_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)
    KeyAscii = Asc(cZeichen)
    
    Select Case Index
        Case Is = 0
            'wegen Volltextsuche nicht mehr gültig
            cValid = "1234567890" & Chr$(8)
        Case Is = 1
            If Label2(7).Caption = "Stück" Then
                If gbNONEGZU Then 'keine negativen
                    cValid = "1234567890+" & Chr$(8)
                Else
                    cValid = "1234567890+-" & Chr$(8)
                End If
            ElseIf Label2(7).Caption = "Kg" Then
                cValid = "1234567890+-," & Chr$(8)
            End If
        Case Is = 2
            cValid = "1234567890," & Chr$(8)
        Case Is = 3
            cValid = "1234567890," & Chr$(8)
        Case Is = 4
            cValid = "1234567890" & Chr$(8)
        Case Is = 5, 8, 9, 12, 11
            cValid = "1234567890" & Chr$(8)
        Case Is = 7, 10
            cValid = gcNUM & gcUPPER & gcLower & Chr(8)
        Case Is = 13
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%"

    End Select
    
    If Index <> 0 And Index <> 6 Then
        If InStr(cValid, cZeichen) = 0 Then
            KeyAscii = 0
        End If
    End If
    
    If Index = 1 And cZeichen = "," Then
        If InStr(Text1(Index).Text, ",") > 0 Then
            KeyAscii = 0
        End If
    End If
    
    If Index = 2 And cZeichen = "," Then
        If InStr(Text1(Index).Text, ",") > 0 Then
            KeyAscii = 0
        End If
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    bscanner = False
    If KeyCode = vbKeyEscape Then
        Form_Unload 1
    End If
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            If gbscanmodi Then
                bscanner = True
            Else
                bscanner = False
            End If
            Command1_Click
        ElseIf Index >= 1 And Index < 7 Then
            Command2_Click 15
        ElseIf Index >= 7 And Index < 13 Then '7-12
            cmdGo_Click
        End If
    End If
    
    If KeyCode = vbKeyF5 Then
        Text1(3).SetFocus
    End If
    
    If KeyCode = vbKeyF4 Then
        If Index = 0 Then
            ctmp = Trim$(Text1(4).Text)
            If ctmp = "" Then
                MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                Text1(4).SetFocus
                Exit Sub
            End If
            
            gF2Prompt.cFeld = "ARTNRPOS"
            gF2Prompt.cWert = ctmp
            gF2Prompt.cWert2 = ""
            gF2Prompt.cWahl = ""
            gF2Prompt.bMultiple = False
            
            Command1_Click
            ctmp = Trim$(Text1(0).Text)
            If ctmp = "" Then
                MsgBox "Bitte den Artikel eindeutig bestimmen (Artikelnummer oder EAN-Code)!", vbCritical, "STOP!"
                Text1(0).SetFocus
                Exit Sub
            End If
            gF2Prompt.cWert2 = ctmp
        
            If gF2Prompt.cFeld <> "" Then
                
                frmWK00a.Show 1
            
                If gF2Prompt.cWahl <> "" Then
                    Text1(Index).Text = gF2Prompt.cWahl
                    If Index = 0 Then
                        Command1_Click
                    End If
                End If
            End If
        End If
    End If
    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False

        Select Case Index
            Case Is = 0     'Artikel
                ctmp = Trim$(Text1(4).Text)
                If ctmp = "" Then
                    MsgBox "Bitte einen Lieferanten angeben!", vbCritical, "STOP!"
                    Text1(4).SetFocus
                    Exit Sub
                Else
                    gF2Prompt.cFeld = "ARTNRPOS"
                    gF2Prompt.cWert = ctmp
                End If
            
            Case Is = 4     'Lieferant
                gF2Prompt.cFeld = "LINR"
        End Select
        
        If gF2Prompt.cFeld <> "" Then
            
            frmWK00a.Show 1
        
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
                If Index = 0 Then
                    Command1_Click
                End If
            End If
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    
    If Index = 4 Then
        ctmp = Text1(4).Text
        ctmp = Trim$(Str$(Val(ctmp)))
        
        cSQL = "Select * from LISRT where LINR = " & ctmp & " "
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!LIEFBEZ) Then
                Label2(4).Caption = rsrs!LIEFBEZ
            Else
                Label2(4).Caption = ""
            End If
        Else
            Label2(4).Caption = ""
        End If
        rsrs.Close: Set rsrs = Nothing
        LeseLieferantenPreisWKL15
    End If
    
    Text1(Index).BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text3_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    
    bscanner = False
    
    If KeyCode = vbKeyEscape Then
        Form_Unload 1
    End If
    
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            If gbscanmodi Then
                bscanner = True
            Else
                bscanner = False
            End If
            
            
            'Mensch wenn du von hier weitergehst, dann auch mal den eventuell geänderten EK speichern
            sNeuerEK = Text1(3).Text
            
            Command1_Click 'suchen
            
            If sNeuerEK <> "" Then
                Text1(3).Text = sNeuerEK
            End If
            
        ElseIf Index >= 1 And Index < 7 Then
            Command2_Click 15
        ElseIf Index >= 7 And Index < 13 Then '7-12
            cmdGo_Click
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(Index).BackColor = vbWhite

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Einzellieferung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
