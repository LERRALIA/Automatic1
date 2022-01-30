VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL171 
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame8 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'Kein
      Height          =   6735
      Left            =   11400
      TabIndex        =   101
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
      Begin VB.TextBox Text1 
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
         Index           =   10
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         Index           =   8
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   1080
         Width           =   1095
      End
      Begin sevCommand3.Command Command5 
         Height          =   345
         Index           =   30
         Left            =   11280
         TabIndex        =   102
         Top             =   240
         Width           =   345
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "?"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   31
         Left            =   9600
         TabIndex        =   103
         Top             =   7680
         Width           =   2055
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   32
         Left            =   5040
         TabIndex        =   113
         Top             =   840
         Width           =   2055
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   2
         Left            =   1800
         TabIndex        =   114
         ToolTipText     =   "Kalender"
         Top             =   960
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   3
         Left            =   4080
         TabIndex        =   115
         ToolTipText     =   "Kalender"
         Top             =   960
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   4
         Left            =   1440
         TabIndex        =   116
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   5
         Left            =   1440
         TabIndex        =   117
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   6
         Left            =   3720
         TabIndex        =   118
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   11
         Left            =   3720
         TabIndex        =   119
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":0000
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   140
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Kunden"
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
         Index           =   9
         Left            =   120
         TabIndex        =   139
         Top             =   6480
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":030A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   138
         Top             =   6000
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Lieferanten"
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
         Index           =   8
         Left            =   120
         TabIndex        =   137
         Top             =   6000
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":0614
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   136
         Top             =   5520
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Artikel"
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
         Index           =   7
         Left            =   120
         TabIndex        =   135
         Top             =   5520
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":091E
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   134
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Zugänge"
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
         Index           =   6
         Left            =   120
         TabIndex        =   133
         Top             =   5040
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":0C28
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   132
         Top             =   4560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Kassenbuch"
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
         Index           =   5
         Left            =   120
         TabIndex        =   131
         Top             =   4560
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":0F32
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   130
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Gutscheine"
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
         Left            =   120
         TabIndex        =   129
         Top             =   4080
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":123C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   128
         Top             =   3600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Bediener"
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
         Index           =   3
         Left            =   120
         TabIndex        =   127
         Top             =   3600
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":1546
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   126
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Ein- und Auszahlungen"
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
         Index           =   2
         Left            =   120
         TabIndex        =   125
         Top             =   3120
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":1850
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   124
         Top             =   2640
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Verkaufsdaten"
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
         Index           =   1
         Left            =   120
         TabIndex        =   123
         Top             =   2640
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "Beschreibung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":1B5A
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   122
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "bis:"
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
         Index           =   40
         Left            =   2400
         TabIndex        =   121
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "von:"
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
         Index           =   39
         Left            =   120
         TabIndex        =   120
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GDPdU - Ausgabe"
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
         TabIndex        =   109
         Top             =   120
         Width           =   5175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   2
         X1              =   120
         X2              =   11640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Anzeige"
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
         Index           =   67
         Left            =   120
         TabIndex        =   108
         Top             =   7800
         Width           =   9255
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "Grundsätze zum Datenzugriff und zur Prüfbarkeit digitaler Unterlagen"
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
         Index           =   66
         Left            =   5160
         TabIndex        =   107
         Top             =   440
         Width           =   6135
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Tagesumsätze"
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
         Index           =   0
         Left            =   120
         TabIndex        =   106
         Top             =   2160
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "hier downloaden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   6000
         MouseIcon       =   "frmWKL171.frx":1E64
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   105
         Top             =   7200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "aktuelles Handbuch"
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
         Index           =   57
         Left            =   120
         TabIndex        =   104
         Top             =   7200
         Width           =   5535
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'Kein
      Height          =   4815
      Left            =   10200
      TabIndex        =   52
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
      Begin VB.ComboBox cboBestHist 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4080
         TabIndex        =   90
         Text            =   "cboBestHist"
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox Text1 
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
         Index           =   6
         Left            =   6240
         TabIndex        =   86
         Top             =   2040
         Width           =   615
      End
      Begin sevCommand3.Command Command5 
         Height          =   345
         Index           =   27
         Left            =   11280
         TabIndex        =   53
         Top             =   240
         Width           =   345
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "?"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   29
         Left            =   9600
         TabIndex        =   54
         Top             =   7680
         Width           =   2055
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   17
         Left            =   6240
         TabIndex        =   60
         Top             =   1560
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "csv"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   18
         Left            =   6960
         TabIndex        =   61
         Top             =   1560
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Beschreibung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   19
         Left            =   6240
         TabIndex        =   71
         Top             =   2520
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "csv"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   20
         Left            =   6960
         TabIndex        =   72
         Top             =   2520
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Beschreibung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   21
         Left            =   6240
         TabIndex        =   76
         Top             =   3000
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "csv"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   22
         Left            =   6960
         TabIndex        =   77
         Top             =   3000
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Beschreibung"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   23
         Left            =   6960
         TabIndex        =   85
         Top             =   2040
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   24
         Left            =   6960
         TabIndex        =   94
         Top             =   4440
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   33
         Left            =   8640
         TabIndex        =   177
         Top             =   4440
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "als Mail versenden"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   37
         Left            =   4080
         MouseIcon       =   "frmWKL171.frx":216E
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   93
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   36
         Left            =   5160
         MouseIcon       =   "frmWKL171.frx":2478
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   92
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "-"
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
         Index           =   35
         Left            =   5040
         TabIndex        =   91
         Top             =   3480
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "verfügbare Bestandshistorien"
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
         Index           =   34
         Left            =   120
         TabIndex        =   89
         Top             =   4440
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "alle Rechnungen"
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
         Index           =   33
         Left            =   120
         TabIndex        =   88
         Top             =   3960
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "alle Kassenbons"
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
         Index           =   32
         Left            =   120
         TabIndex        =   87
         Top             =   3480
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "-"
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
         Index           =   31
         Left            =   5040
         TabIndex        =   84
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   30
         Left            =   4080
         MouseIcon       =   "frmWKL171.frx":2782
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   83
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   29
         Left            =   5160
         MouseIcon       =   "frmWKL171.frx":2A8C
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   82
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "-"
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
         Index           =   28
         Left            =   5040
         TabIndex        =   81
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "alle Bestandsänderungen"
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
         Index           =   27
         Left            =   120
         TabIndex        =   80
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   26
         Left            =   5160
         MouseIcon       =   "frmWKL171.frx":2D96
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   79
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   25
         Left            =   4080
         MouseIcon       =   "frmWKL171.frx":30A0
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   78
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "-"
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
         Index           =   24
         Left            =   5040
         TabIndex        =   75
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   23
         Left            =   4080
         MouseIcon       =   "frmWKL171.frx":33AA
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   74
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   22
         Left            =   5160
         MouseIcon       =   "frmWKL171.frx":36B4
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   73
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "alle Preisänderungen"
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
         Index           =   21
         Left            =   120
         TabIndex        =   70
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "alle Kassenabrechnungen"
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
         Index           =   20
         Left            =   120
         TabIndex        =   69
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "aktuelles Handbuch"
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
         Index           =   19
         Left            =   120
         TabIndex        =   68
         Top             =   4920
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
         Caption         =   "hier downloaden"
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
         Left            =   4080
         MouseIcon       =   "frmWKL171.frx":39BE
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   67
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   17
         Left            =   5160
         MouseIcon       =   "frmWKL171.frx":3CC8
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   66
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00C0C000&
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
         Index           =   16
         Left            =   4080
         MouseIcon       =   "frmWKL171.frx":3FD2
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   65
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "-"
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
         Index           =   14
         Left            =   5040
         TabIndex        =   64
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "Pfadangabe"
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
         Index           =   13
         Left            =   4800
         TabIndex        =   63
         Top             =   840
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H008080FF&
         Caption         =   "Speicherort der Datenbank:"
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
         Left            =   2160
         TabIndex        =   62
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "alle Buchungssätze/Verkaufsvorgänge"
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
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "Stand:"
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
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H008080FF&
         Caption         =   "Grundsätze zum Datenzugriff und zur Prüfbarkeit digitaler Unterlagen"
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
         TabIndex        =   57
         Top             =   440
         Width           =   7335
      End
      Begin VB.Label Label1 
         Caption         =   "Anzeige"
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
         Index           =   15
         Left            =   120
         TabIndex        =   56
         Top             =   7800
         Width           =   9255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   120
         X2              =   11640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "GDPdU"
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
         TabIndex        =   55
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Height          =   1215
      Left            =   10080
      TabIndex        =   46
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "GDPdU - Komplettausgabe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   110
         Top             =   1680
         Width           =   4455
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "DATEV-Ausgabe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   48
         Top             =   3120
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "GDPdU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   47
         Top             =   2400
         Width           =   4455
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   16
         Left            =   9600
         TabIndex        =   49
         Top             =   7080
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   15
         Left            =   9600
         TabIndex        =   50
         Top             =   7680
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "um an die gewünschten Daten bzw. Informationen zu kommen, gehen Sie bitte wie folgt vor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Index           =   44
         Left            =   240
         TabIndex        =   176
         Top             =   3960
         Width           =   11415
      End
      Begin VB.Label Label6 
         Caption         =   "Welchen Programmteil möchten Sie öffnen?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   10935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'Kein
      Height          =   14295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   25815
      Begin VB.Frame Frame7 
         Caption         =   "Einstellungen"
         Height          =   3615
         Left            =   1200
         TabIndex        =   95
         Top             =   6120
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CheckBox Check1 
            Caption         =   "DATEV-Format mit Festschreibekennzeichen (ab Version 5.1)"
            Height          =   375
            Left            =   120
            TabIndex        =   178
            Top             =   960
            Width           =   5415
         End
         Begin VB.TextBox Text1 
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
            Index           =   7
            Left            =   120
            MaxLength       =   10
            TabIndex        =   96
            Top             =   600
            Width           =   1335
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   26
            Left            =   4320
            TabIndex        =   97
            Top             =   240
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
            Caption         =   "Speichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   28
            Left            =   6120
            TabIndex        =   98
            Top             =   240
            Width           =   375
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
            Caption         =   "x"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Ausgabe Datumsformat"
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
            Index           =   38
            Left            =   120
            TabIndex        =   99
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.TextBox Text1 
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
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Filialkonten/Kostenstellen"
         Height          =   735
         Left            =   0
         TabIndex        =   22
         Top             =   4440
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox Text1 
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   27
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox Text1 
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
            Index           =   1
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   6375
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   10
            Left            =   4320
            TabIndex        =   23
            Top             =   720
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
            Caption         =   "Löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   7
            Left            =   4320
            TabIndex        =   24
            Top             =   240
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
            Caption         =   "Speichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   4
            Left            =   6120
            TabIndex        =   28
            Top             =   240
            Width           =   375
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
            Caption         =   "x"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kostenstelle"
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
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Filialkonto"
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
            Left            =   1560
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datev Konten"
         Height          =   2295
         Left            =   3720
         TabIndex        =   11
         Top             =   4680
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox Text1 
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
            Index           =   2
            Left            =   120
            MaxLength       =   4
            TabIndex        =   16
            Top             =   480
            Width           =   975
         End
         Begin VB.ComboBox Combo3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            Style           =   2  'Dropdown-Liste
            TabIndex        =   15
            Top             =   480
            Width           =   3015
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4680
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   7695
         End
         Begin sevCommand3.Command Command5 
            Height          =   495
            Index           =   14
            Left            =   9480
            TabIndex        =   12
            Top             =   5520
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
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
            Caption         =   "benutzerdefiniertes Konto"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   9
            Left            =   6120
            TabIndex        =   13
            Top             =   480
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
            Caption         =   "Löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   8
            Left            =   4320
            TabIndex        =   17
            Top             =   480
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
            Caption         =   "Speichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   495
            Index           =   6
            Left            =   9480
            TabIndex        =   18
            Top             =   6120
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
            Caption         =   "Zurück"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Konto"
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
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kontenbezeichnung:"
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
            Left            =   1200
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
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
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Benutzerdefiniertes Konto erstellen"
         Height          =   1575
         Left            =   4320
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   9
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   164
            Top             =   4440
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   8
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   163
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   7
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   162
            Top             =   3720
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   6
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   161
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   5
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   160
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   4
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   159
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   3
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   158
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   2
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   157
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   1
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   156
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Index           =   0
            Left            =   6360
            MaxLength       =   6
            TabIndex        =   155
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   0
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   150
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   1
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   149
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   2
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   148
            Top             =   1920
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   3
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   147
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   4
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   146
            Top             =   2640
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   5
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   145
            Top             =   3000
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   6
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   144
            Top             =   3360
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   7
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   143
            Top             =   3720
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   8
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   142
            Top             =   4080
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Index           =   9
            Left            =   7320
            MaxLength       =   35
            TabIndex        =   141
            Top             =   4440
            Width           =   2055
         End
         Begin VB.ListBox List3 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3360
            Left            =   120
            TabIndex        =   6
            Top             =   1920
            Width           =   5895
         End
         Begin VB.TextBox Text1 
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
            Index           =   4
            Left            =   120
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1440
            Width           =   975
         End
         Begin VB.TextBox Text1 
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
            Index           =   5
            Left            =   1200
            MaxLength       =   35
            TabIndex        =   2
            Top             =   1440
            Width           =   2895
         End
         Begin sevCommand3.Command Command5 
            Height          =   495
            Index           =   13
            Left            =   9480
            TabIndex        =   3
            Top             =   6120
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
            Caption         =   "Zurück"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   495
            Index           =   12
            Left            =   9480
            TabIndex        =   4
            Top             =   5520
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
            Caption         =   "Speichern"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin sevCommand3.Command Command5 
            Height          =   375
            Index           =   11
            Left            =   4320
            TabIndex        =   7
            Top             =   1440
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
            Caption         =   "Löschen"
            PictureAlign    =   2
            PictureVisible  =   0   'False
            Version3        =   -1  'True
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   175
            Top             =   4440
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   174
            Top             =   4080
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   173
            Top             =   3720
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   172
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   171
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   170
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   169
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   168
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   167
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C000&
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
            Left            =   9600
            TabIndex        =   166
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kontenbezeichnung"
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
            Index           =   43
            Left            =   7320
            TabIndex        =   165
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "KontoNr"
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
            Index           =   42
            Left            =   6360
            TabIndex        =   154
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Auszahlungsgrund"
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
            Index           =   41
            Left            =   9600
            TabIndex        =   153
            Top             =   840
            Width           =   1935
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808000&
            BorderWidth     =   2
            Index           =   3
            X1              =   6120
            X2              =   6120
            Y1              =   240
            Y2              =   6120
         End
         Begin VB.Label lbl6 
            Caption         =   "Konten für bestimmte Artikelgruppen"
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
            TabIndex        =   152
            Top             =   360
            Width           =   4095
         End
         Begin VB.Label lbl6 
            Caption         =   "Konten für bestimmte Auszahlungsvorgänge"
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
            Index           =   24
            Left            =   6360
            TabIndex        =   151
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C000&
            Caption         =   "Kontenbezeichnung:"
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
            Left            =   1200
            TabIndex        =   9
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label1 
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
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   975
         End
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   5
         Left            =   8160
         TabIndex        =   21
         Top             =   840
         Width           =   1695
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
         Caption         =   "Datev Konten"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   3
         Left            =   9960
         TabIndex        =   31
         Top             =   840
         Width           =   1695
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
         Caption         =   "Filialkonten / Kostenstellen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   345
         Index           =   2
         Left            =   11280
         TabIndex        =   33
         Top             =   240
         Width           =   345
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6
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
         Caption         =   "?"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   1
         Left            =   5040
         TabIndex        =   34
         Top             =   840
         Width           =   2055
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
         Caption         =   "Export"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   35
         Top             =   7680
         Width           =   2055
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   1
         Left            =   1800
         TabIndex        =   36
         ToolTipText     =   "Kalender"
         Top             =   960
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
      Begin sevCommand3.Command Command0 
         Height          =   405
         Index           =   0
         Left            =   4080
         TabIndex        =   37
         ToolTipText     =   "Kalender"
         Top             =   960
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   9
         Left            =   1440
         TabIndex        =   38
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   10
         Left            =   1440
         TabIndex        =   39
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   7
         Left            =   3720
         TabIndex        =   40
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Zurück"
         ToolTipTitle    =   "Zurück"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command0 
         Height          =   165
         Index           =   8
         Left            =   3720
         TabIndex        =   41
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   291
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
         ToolTip         =   "Vor"
         ToolTipTitle    =   "Vor"
         ButtonStyle     =   2
         Caption         =   ""
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   25
         Left            =   7560
         TabIndex        =   100
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
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
         Caption         =   "E"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label lblUeberschrift 
         BackStyle       =   0  'Transparent
         Caption         =   "DATEV Konten"
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
         TabIndex        =   45
         Top             =   120
         Width           =   6135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   11640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Anzeige"
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
         Index           =   4
         Left            =   120
         TabIndex        =   44
         Top             =   7800
         Width           =   9255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "von:"
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
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "bis:"
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
         Left            =   2400
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmWKL171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command0_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lDat As Long

    Select Case index
        Case Is = 1        ' Kalender
            Text1(9).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 0        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
         Case Is = 2        ' Kalender
            Text1(8).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
        Case Is = 3        ' Kalender
            Text1(10).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
        Case 5
            If IsDate(Text1(8).Text) = False Then
                Text1(8).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(8).Text) = True Then
                    lDat = CLng(DateValue(Text1(8).Text))
                End If
                lDat = lDat + 1
                Text1(8).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 4
            If IsDate(Text1(8).Text) = False Then
                Text1(8).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(8).Text) = True Then
                    lDat = CLng(DateValue(Text1(8).Text))
                End If
                lDat = lDat - 1
                Text1(8).Text = Format(lDat, "DD.MM.YYYY")
            End If
            
            
            
        Case 10
            If IsDate(Text1(9).Text) = False Then
                Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(9).Text) = True Then
                    lDat = CLng(DateValue(Text1(9).Text))
                End If
                lDat = lDat + 1
                Text1(9).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 9
            If IsDate(Text1(9).Text) = False Then
                Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(9).Text) = True Then
                    lDat = CLng(DateValue(Text1(9).Text))
                End If
                lDat = lDat - 1
                Text1(9).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 8
            If IsDate(Text1(3).Text) = False Then
                Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(3).Text) = True Then
                    lDat = CLng(DateValue(Text1(3).Text))
                End If
                lDat = lDat + 1
                Text1(3).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 7
            If IsDate(Text1(3).Text) = False Then
                Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(3).Text) = True Then
                    lDat = CLng(DateValue(Text1(3).Text))
                End If
                lDat = lDat - 1
                Text1(3).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 11
            If IsDate(Text1(10).Text) = False Then
                Text1(10).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(10).Text) = True Then
                    lDat = CLng(DateValue(Text1(10).Text))
                End If
                lDat = lDat + 1
                Text1(10).Text = Format(lDat, "DD.MM.YYYY")
            End If
        Case 6
            If IsDate(Text1(10).Text) = False Then
                Text1(10).Text = Format(DateValue(Now), "DD.MM.YYYY")
            Else
                If IsDate(Text1(10).Text) = True Then
                    lDat = CLng(DateValue(Text1(10).Text))
                End If
                lDat = lDat - 1
                Text1(10).Text = Format(lDat, "DD.MM.YYYY")
            End If
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command5_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lBis            As Long
    Dim lVon            As Long
    Dim sdat            As String
    Dim lDay            As Long
    Dim iRet            As Integer

    Select Case index
        Case 0
            Frame1.Visible = False
            Frame5.Visible = True
        Case 1 'Export
            Screen.MousePointer = 11
            
            lVon = CLng(DateValue(Text1(9).Text))
            lBis = CLng(DateValue(Text1(3).Text))
            
            loeschNEW "DATEVEXPORT", gdBase
            CreateTableT2 "DATEVEXPORT", gdBase
            
            For lDay = lVon To lBis
                anzeige "normal", Format(lDay, "DD.MM.YY"), Label1(4)
                EXPORT lDay
            Next lDay
            
            sdat = Format$(Text1(9).Text, "DDMM") & Format$(Text1(3).Text, "DDMM")
            
            Screen.MousePointer = 0
            
            If Check1.value = vbChecked Then
                ExportCSV_Bauch sdat
            Else
                ExportCSV sdat
            End If

        Case 2
            gsHelpstring = "DATEV Konten"
            frmWKL110.Show 1
        Case 3
            Frame4.Visible = True
            Frame2.Visible = False
            Frame7.Visible = False
            ZeigeFilKonten
        Case 4
            Frame4.Visible = False
        Case 5
            Frame2.Visible = True
            Frame4.Visible = False
            Frame7.Visible = False
            fülleDATEVALLG Combo3
            ZeigeKonten
        Case 6
            Frame2.Visible = False
        Case 7
            SpeicherFilKonten Text1(0).Text, Text1(1).Text
            ZeigeFilKonten
        Case 8
            SpeicherKonten Val(Text1(2).Text), Combo3.Text
            ZeigeKonten
        Case 9
            LoescheKonten
            ZeigeKonten
        Case 10
            LoescheFilKonten
            ZeigeFilKonten
        Case 11
            LoescheKontenAGN
            ZeigeKontenAGN
        Case 12
            speicherAuszahlungsgrund
            ZeigeKontenAuszahlung
            
            SpeicherKontenAGN Val(Text1(4).Text), Text1(5)
            ZeigeKontenAGN
        Case 13
            Frame3.Visible = False
            Frame2.Visible = True
            fülleDATEVALLG Combo3
        Case 14
            Frame3.Visible = True
            Frame2.Visible = False
            ZeigeKontenAGN
            ZeigeKontenAuszahlung
            
        Case 15
            Unload frmWKL171
        Case 16
            Zeigeauswahlframe
        Case 17
            GDPdU_ExportCSV_Kassjour Label1(16), Label1(17)
        Case 18
            Beschreibung_GDPdU_ExportCSV_Kassjour
            zeigeHilfeDabapfad "GDPdU", "Beschreibung_Verkäufe_xxx_xxx.txt"
        Case 19
            GDPdU_ExportCSV_KVKPR1PROT Label1(23), Label1(22)
        Case 20
            Beschreibung_GDPdU_ExportCSV_KVKPR1PROT
            zeigeHilfeDabapfad "GDPdU", "Beschreibung_Preisänderungen_xxx_xxx.txt"
        Case 21
            GDPdU_ExportCSV_BESTPROT Label1(25), Label1(26)
        Case 22
            Beschreibung_GDPdU_ExportCSV_BESTPROT
            zeigeHilfeDabapfad "GDPdU", "Beschreibung_Bestandsänderungen_xxx_xxx.txt"
        Case 23
            If Text1(6).Text <> "" Then
                drucke_nachträglich_Zbon CLng(Text1(6).Text)
            End If
        Case 24
            If cboBestHist.Text <> "" Then
                anzeige "normal", "", Label1(15)
                Screen.MousePointer = 11
                
                iRet = MsgBox("Möchten Sie die Werte als Lieferantenübersicht angezeigt bekommen? (Nein = Artikelansicht)", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbYes Then
                    zeige_Best_Hist_GDPdU Left(cboBestHist.Text, 8), Label1(15)
                Else
                    zeige_Best_Hist_Einzel_GDPdU Left(cboBestHist.Text, 8), Label1(15)
                End If
                
                Screen.MousePointer = 0
            Else
                anzeige "rot", "Wählen Sie bitte ein Datum aus!", Label1(15)
            End If
        Case 33 'als Mail
            If cboBestHist.Text <> "" Then
                anzeige "normal", "", Label1(15)
                Screen.MousePointer = 11
                
                iRet = MsgBox("Möchten Sie die Werte als Lieferantenübersicht angezeigt bekommen? (Nein = Artikelansicht)", vbQuestion + vbYesNo, "Winkiss Frage:")
                If iRet = vbYes Then
                    zeige_Best_Hist_GDPdU Left(cboBestHist.Text, 8), Label1(15), True
                Else
                    zeige_Best_Hist_Einzel_GDPdU Left(cboBestHist.Text, 8), Label1(15), True
                End If
                
                Screen.MousePointer = 0
            Else
                anzeige "rot", "Wählen Sie bitte ein Datum aus!", Label1(15)
            End If
        
        
        Case 25
            Frame2.Visible = False
            Frame4.Visible = False
            Frame7.Visible = True
        Case 26
            voreinstellungspeichernE171
        Case 27
            gsHelpstring = "GDPdU"
            frmWKL110.Show 1
        Case 28
            Frame7.Visible = False
        Case 29
            Frame6.Visible = False
            Frame5.Visible = True
            
        Case 31
            Frame8.Visible = False
            Frame5.Visible = True
        Case 30
            gsHelpstring = "GDPdU - Ausgabe"
            frmWKL110.Show 1
        Case 32
            GDPDU_Export Text1(8).Text, Text1(10).Text
        Case 34
'            Datev_ExportCSV_Kassjour Text1(9).Text, Text1(3).Text, Text1(11).Text
    End Select
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub GDPDU_Export(sVon As String, sBis As String)
    On Error GoTo LOKAL_ERROR
    
    Dim db As Database
    Dim cPfad As String
    Dim sDatenbankname As String
    
    Dim lVon As Long
    Dim lBis As Long
    Dim sSQL As String
    Dim sVonBez As String
    Dim sBisBez As String
    
    
    Dim cPfad2 As String
    cPfad2 = gcDBPfad
    If Right$(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    sVonBez = SwapStr(sVon, ".", "")
    sBisBez = SwapStr(sBis, ".", "")
    
    lVon = CLng(DateValue(sVon))
    lBis = CLng(DateValue(sBis))
    
    sDatenbankname = "GDPDU_" & sVonBez & "_" & sBisBez & ".mdb"
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    Kill cPfad & sDatenbankname
    Set db = CreateDatabase(cPfad & sDatenbankname, dbLangGeneral, dbVersion40)
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Die Daten für den erforderlichen Überprüfungszeitraum werden ermittelt...", Label1(67)
    
    'Umsatz

    Label33(0).Visible = True
    anzeige "normal", "Tagesumsätze werden exportiert...", Label33(0)

    sSQL = "Create Table Tagesumsaetze_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Datum Datetime "
    sSQL = sSQL & " ,UmsatzvolleSteuer Double "
    sSQL = sSQL & " ,UmsatzermSteuer Double "
    sSQL = sSQL & " ,UmsatzohneSteuer Double "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Tagesumsaetze_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " Select "
    sSQL = sSQL & " Datum "
    sSQL = sSQL & " , umsv1 as UmsatzvolleSteuer  "
    sSQL = sSQL & " , umse1 as UmsatzermSteuer  "
    sSQL = sSQL & " , umso1 as UmsatzohneSteuer  "
    
    sSQL = sSQL & "  from UMSATZ in '" & cPfad2 & "kissdata.MDB'"
    sSQL = sSQL & " where Datum between " & lVon & " and " & lBis
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Tagesumsätze, Fertig!", Label33(0)
    Label4(0).Visible = True
    
    
    'Kassjour

    Label33(1).Visible = True
    anzeige "normal", "Verkaufsdaten werden exportiert...", Label33(1)

    sSQL = "Create Table Verkaeufe_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Verkaufsdatum Datetime "
    sSQL = sSQL & " ,Verkaufszeit Text(10)"
    sSQL = sSQL & " ,Kassennummer BYTE "
    sSQL = sSQL & " ,Artikelnummer LONG "
    sSQL = sSQL & " ,Artikelbezeichnung Text(35) "
    sSQL = sSQL & " ,Verkaufsmenge LONG "
    sSQL = sSQL & " ,Kassenverkaufspreis Double "
    sSQL = sSQL & " ,Verkaufspreis Double "
    sSQL = sSQL & " ,Kundennummer LONG "
    sSQL = sSQL & " ,MWST Text(1) "
    sSQL = sSQL & " ,umsatzsteuerpflichtig Text(1) "
    sSQL = sSQL & " ,Zahlungsart Text(2) "
    sSQL = sSQL & " ,Kassenbonnr integer "
    sSQL = sSQL & " ,Bediener integer "
    sSQL = sSQL & " ,Filiale byte "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Verkaeufe_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " Select "
    sSQL = sSQL & " adate as Verkaufsdatum "
    sSQL = sSQL & " , azeit as Verkaufszeit "
    sSQL = sSQL & " , kasnum as Kassennummer "
    sSQL = sSQL & " , Artnr as Artikelnummer "
    sSQL = sSQL & " , bezeich as Artikelbezeichnung "
    sSQL = sSQL & " , Menge as Verkaufsmenge "
    sSQL = sSQL & " , preis as Kassenverkaufspreis "
    sSQL = sSQL & " , vkpr as Verkaufspreis "
    sSQL = sSQL & " , kundnr as Kundennummer "
    sSQL = sSQL & " , MWST "
    sSQL = sSQL & " , UMS_OK as umsatzsteuerpflichtig "
    sSQL = sSQL & " , KK_ART as Zahlungsart "
    sSQL = sSQL & " , Belegnr as Kassenbonnr "
    sSQL = sSQL & " , Bediener  "
    sSQL = sSQL & " , Filiale "
    sSQL = sSQL & "  from Kassjour in '" & cPfad2 & "kissdata.MDB'"
    sSQL = sSQL & " where adate between " & lVon & " and " & lBis
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Verkaufsdaten, Fertig!", Label33(1)
    Label4(1).Visible = True
    
    
    
    
    
    
    'Ein und Auszahlungen

    Label33(2).Visible = True
    anzeige "normal", "Ein- und Auszahlungen werden exportiert...", Label33(2)
    
    sSQL = "Create Table EinAuszahlungen_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Datum Datetime "
    sSQL = sSQL & " ,Zeit Text(10)"
    sSQL = sSQL & " ,Kassennummer BYTE "
    sSQL = sSQL & " ,Bediener INTEGER "
    sSQL = sSQL & " ,Art Text(20) "
    sSQL = sSQL & " ,Betrag Double "
    sSQL = sSQL & " ,Verwendungszweck Text(50)"
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into EinAuszahlungen_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " Select "
    sSQL = sSQL & " adate as Datum "
    sSQL = sSQL & " , azeit as ZEIT "
    sSQL = sSQL & " , Kasnum as Kassennummer  "
    sSQL = sSQL & " , BEDNU as Bediener  "
    sSQL = sSQL & " , Art "
    sSQL = sSQL & " , Betrag "
    sSQL = sSQL & " ,bezeich as Verwendungszweck "
    
    sSQL = sSQL & "  from KAEINAUSF in '" & cPfad2 & "kissdata.MDB'"
    sSQL = sSQL & " where adate between " & lVon & " and " & lBis
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Ein- und Auszahlungen, Fertig!", Label33(2)
    Label4(2).Visible = True
    

    'Bediener
    Label33(3).Visible = True
    anzeige "normal", "Bediener werden exportiert...", Label33(3)

    sSQL = "Create Table Bediener_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Bedienernummer integer "
    sSQL = sSQL & " ,Bedienerbezeichnung Text(40) "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Bediener_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " Select "
    sSQL = sSQL & " Bednu as Bedienernummer  "
    sSQL = sSQL & " ,Bedname as Bedienerbezeichnung  "
    sSQL = sSQL & "  from BEDNAME in '" & cPfad2 & "kissdata.MDB'"
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Bediener, Fertig!", Label33(3)
    Label4(3).Visible = True
    

    'Gutscheine
    Label33(4).Visible = True
    anzeige "normal", "Gutscheine werden exportiert...", Label33(4)

    sSQL = "Create Table Gutscheine_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Gutscheinnummer integer "
    sSQL = sSQL & " ,Gutscheinwert double "
    sSQL = sSQL & " ,Gutschein_AusgabeDatum DATETIME "
    sSQL = sSQL & " ,Gutschein_EinloeseDatum DATETIME "
    sSQL = sSQL & " ,Gutschein_Ausgabefiliale integer "
    sSQL = sSQL & " ,Gutschein_AusgabeBediener integer "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Gutscheine_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " Select "
    sSQL = sSQL & " Gutschnr as Gutscheinnummer  "
    sSQL = sSQL & " ,Wert as Gutscheinwert  "
    sSQL = sSQL & " ,Dat_ausg as Gutschein_AusgabeDatum  "
    sSQL = sSQL & " ,Dat_einl as Gutschein_EinloeseDatum  "
    sSQL = sSQL & " ,Filiale as Gutschein_Ausgabefiliale  "
    sSQL = sSQL & " ,Bednu as Gutschein_AusgabeBediener  "
    sSQL = sSQL & "  from Gutsch in '" & cPfad2 & "kissdata.MDB'"
    sSQL = sSQL & " where Dat_ausg between " & lVon & " and " & lBis
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Gutscheine, Fertig!", Label33(4)
    Label4(4).Visible = True
    
    'Kassenbuch

    Label33(5).Visible = True
    anzeige "normal", "Kassenbuch wird exportiert...", Label33(5)

    sSQL = "Create Table Kassenbuch_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Datum DATETIME "
    sSQL = sSQL & " ,Position BYTE "
    sSQL = sSQL & " ,Bezeichnung Text(150) "
    sSQL = sSQL & " ,EuroUmsatz double"
    sSQL = sSQL & " ,EuroBar double"
    sSQL = sSQL & " ,Kassenbuch_Kassennummer integer "
    sSQL = sSQL & " ,Kassenbuch_Filiale integer "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Kassenbuch_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " Select "
    sSQL = sSQL & " Datum  "
    sSQL = sSQL & " ,POS as Position "
    sSQL = sSQL & " ,BEZUMS as Bezeichnung "
    sSQL = sSQL & " ,eurums as EuroUmsatz "
    sSQL = sSQL & " ,eurbar as EuroBar "
    sSQL = sSQL & " ,Kasnum as Kassenbuch_Kassennummer  "
    sSQL = sSQL & " ,Filiale as Kassenbuch_Filiale  "
    sSQL = sSQL & "  from KABUCH in '" & cPfad2 & "kissdata.MDB'"
    sSQL = sSQL & " where DATUM between " & lVon & " and " & lBis
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kassenbuch, Fertig!", Label33(5)
    Label4(5).Visible = True
    
    
    
    'Zugänge

    Label33(6).Visible = True
    anzeige "normal", "Zugänge werden exportiert...", Label33(6)

    loeschNEW "Zugang_" & sVonBez & "_" & sBisBez, db
    
    sSQL = "Create Table Zugang_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Zugangsdatum Datetime "
    sSQL = sSQL & " ,Zugangszeit Text(10)"
    sSQL = sSQL & " ,Artikelnummer LONG "
    sSQL = sSQL & " ,Artikelbezeichnung Text(35) "
    sSQL = sSQL & " ,ArtikelEAN Text(13) "
    sSQL = sSQL & " ,Bestand_Alt LONG "
    sSQL = sSQL & " ,Zugangsmenge LONG "
    sSQL = sSQL & " ,Bestand_NEU LONG "
    sSQL = sSQL & " ,Filiale byte "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Zugang_" & sVonBez & "_" & sBisBez
    sSQL = sSQL & " Select "
    sSQL = sSQL & " adate as Zugangsdatum  "
    sSQL = sSQL & " ,uhrzeit as Zugangszeit "
    sSQL = sSQL & " ,artnr as Artikelnummer  "
    sSQL = sSQL & " ,bezeich as Artikelbezeichnung  "
    sSQL = sSQL & " ,ean as ArtikelEAN "
    sSQL = sSQL & " ,bestandalt as Bestand_Alt  "
    sSQL = sSQL & " ,bewegung as Zugangsmenge  "
    sSQL = sSQL & " ,bestandneu as Bestand_NEU  "
    sSQL = sSQL & " ,filialnr as Filiale  "
    sSQL = sSQL & "  from Zugang in '" & cPfad2 & "kissdata.MDB'"
    sSQL = sSQL & " where ADATE between " & lVon & " and " & lBis
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Zugang, Fertig!", Label33(6)
    Label4(6).Visible = True
    
    
    
    
    
    'Artikel

    Label33(7).Visible = True
    anzeige "normal", "Artikel werden exportiert...", Label33(7)

    loeschNEW "Artikel", db
    
    sSQL = "Create Table Artikel"
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Artikelnummer LONG "
    sSQL = sSQL & " ,Artikelbezeichnung Text(35) "
    sSQL = sSQL & " ,ArtikelEAN1 Text(13) "
    sSQL = sSQL & " ,ArtikelEAN2 Text(13) "
    sSQL = sSQL & " ,ArtikelEAN3 Text(13) "
    sSQL = sSQL & " ,Kassenverkaufspreis double "
    sSQL = sSQL & " ,Listenverkaufspreis double "
    sSQL = sSQL & " ,Schnitteinkaufspreis double "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Artikel "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artnr as Artikelnummer  "
    sSQL = sSQL & " ,bezeich as Artikelbezeichnung  "
    sSQL = sSQL & " ,ean as ArtikelEAN1 "
    sSQL = sSQL & " ,ean2 as ArtikelEAN2 "
    sSQL = sSQL & " ,ean3 as ArtikelEAN3 "
    sSQL = sSQL & " ,KVKPR1 as Kassenverkaufspreis "
    sSQL = sSQL & " ,VKPR as Listenverkaufspreis "
    sSQL = sSQL & " ,EKPR as Schnitteinkaufspreis "
    sSQL = sSQL & "  from Artikel in '" & cPfad2 & "kissdata.MDB'"
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Artikel, Fertig!", Label33(7)
    Label4(7).Visible = True
    
    'Lieferanten

    Label33(8).Visible = True
    anzeige "normal", "Lieferanten werden exportiert...", Label33(8)

    loeschNEW "Lieferanten", db
    
    sSQL = "Create Table Lieferanten"
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Lieferantennummer LONG "
    sSQL = sSQL & " ,Lieferantenbezeichnung Text(35) "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Lieferanten "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " linr as Lieferantennummer  "
    sSQL = sSQL & " ,liefbez as Lieferantenbezeichnung  "
    sSQL = sSQL & "  from LISRT in '" & cPfad2 & "kissdata.MDB'"
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Lieferanten, Fertig!", Label33(8)
    Label4(8).Visible = True
    
    
    'Kunden

    Label33(9).Visible = True
    anzeige "normal", "Kunden werden exportiert...", Label33(9)
    
    sSQL = "Create Table Kunden"
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Kundennummer LONG "
    sSQL = sSQL & " ,Kundenname Text(35) "
    sSQL = sSQL & " ,Kundenvorname Text(35) "
    sSQL = sSQL & " ) "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Kunden "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " Kundnr as Kundennummer  "
    sSQL = sSQL & " ,name as Kundenname  "
    sSQL = sSQL & " ,vorname as Kundenvorname  "
    sSQL = sSQL & "  from Kunden in '" & cPfad2 & "kissdata.MDB'"
    db.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Kunden, Fertig!", Label33(9)
    Label4(9).Visible = True
    
    
    
    
    
    anzeige "normal", "Die GDPdU- Ausgabe ist fertiggestellt.", Label1(67)
    MsgBox "Die Datei '" & sDatenbankname & "' befindet sich unter: " & cPfad, vbInformation, "Zentrale Hinweis:"
    
    
    
    Screen.MousePointer = 0
            
     
    db.Close
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "GDPDU_Export"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Private Sub Zeigeauswahlframe()
    On Error GoTo LOKAL_ERROR
    
    Frame5.Visible = False
    
    If Option2(0).value = True Then         'GDPdU
        Frame6.Visible = True
        Aufbereitung_derZahlen
    ElseIf Option2(2).value = True Then     'DATEV
        Frame1.Visible = True
    ElseIf Option2(1).value = True Then
        Frame8.Visible = True 'GDPdU Komplettausgabe
        
        Text1(8).Text = Format("01.01." & Year(DateValue(Now)), "DD.MM.YYYY")
        Text1(10).Text = Format(DateValue(Now), "DD.MM.YYYY")
        
        anzeige "normal", "Geben Sie bitte den Überprüfungszeitraum an!", Label1(67)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Aufbereitung_derZahlen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim GDPdU_DB    As Database
    Dim rsrs        As DAO.Recordset
    Dim dateStand   As Date
    Dim lZbonNr     As Long
    Dim lAnz        As Long
    Dim lDiff       As Long
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    
    
    '1. Stand - ermitteln
    Set rsrs = GDPdU_DB.OpenRecordset("select DATUM from GDPDU_STAND")
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Datum) Then
            dateStand = rsrs!Datum
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    anzeige "normal", "Stand: " & dateStand, Label1(9)
    
    anzeige "normal", cPfad & " (passwortgeschützt)", Label1(13)
    
    'Kassjour
    If NewTableSuchenDBKombi("Kassjour", GDPdU_DB) Then
    
        CheckIndex "Kassjour", "ADATE", "", GDPdU_DB
        'min Kassjourdat
        Set rsrs = GDPdU_DB.OpenRecordset("select min(adate) as Datum from Kassjour")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                dateStand = rsrs!Datum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(16)
        
        'max Kassjourdat
        Set rsrs = GDPdU_DB.OpenRecordset("select max(adate) as Datum from Kassjour")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                dateStand = rsrs!Datum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(17)
        
        Command5(17).Enabled = True
    Else
        Command5(17).Enabled = False
    End If
    
    'Preisänderungen
    If NewTableSuchenDBKombi("KVKPR1PROT", GDPdU_DB) Then
    
        CheckIndex "KVKPR1PROT", "lastdate", "", GDPdU_DB
        'min dat
        Set rsrs = GDPdU_DB.OpenRecordset("select min(lastdate) as Datum from KVKPR1PROT")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                dateStand = rsrs!Datum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(23)
        
        'max dat
        Set rsrs = GDPdU_DB.OpenRecordset("select max(lastdate) as Datum from KVKPR1PROT")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                dateStand = rsrs!Datum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(22)
        
        Command5(19).Enabled = True
    Else
        Command5(19).Enabled = False
    End If
    
    'Bestandsveränderungen
    If NewTableSuchenDBKombi("BESTPROT", GDPdU_DB) Then
    
        CheckIndex "BESTPROT", "lastdate", "", GDPdU_DB
        'min dat
        Set rsrs = GDPdU_DB.OpenRecordset("select min(lastdate) as Datum from BESTPROT")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                dateStand = rsrs!Datum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(25)
        
        'max dat
        Set rsrs = GDPdU_DB.OpenRecordset("select max(lastdate) as Datum from BESTPROT")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                dateStand = rsrs!Datum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(26)
        
        Command5(21).Enabled = True
    Else
        Command5(21).Enabled = False
    End If
    
    'alle Abschlüsse TAGKOPF_TEMP
    If NewTableSuchenDBKombi("TAGKOPF_TEMP", GDPdU_DB) Then
    
        CheckIndex "TAGKOPF_TEMP", "neueANR", "", GDPdU_DB
        'max Zbon Nr
        Set rsrs = GDPdU_DB.OpenRecordset("select max(neueANR) as ZbonNr from TAGKOPF_TEMP")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!ZbonNr) Then
                lZbonNr = rsrs!ZbonNr
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", "Nr. " & CStr(lZbonNr), Label1(29)
        Text1(6).Text = lZbonNr
        
        lDiff = lZbonNr
    
    
        'min Zbon Nr
        Set rsrs = GDPdU_DB.OpenRecordset("select min(neueANR) as ZbonNr from TAGKOPF_TEMP")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!ZbonNr) Then
                lZbonNr = rsrs!ZbonNr
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", "Nr. " & CStr(lZbonNr), Label1(30)
        
        lDiff = lDiff - lZbonNr + 1
        
        
        
        
        'anzahl
        lAnz = 0
        Set rsrs = GDPdU_DB.OpenRecordset("select count(*) as maxi from TAGKOPF_TEMP")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!maxi) Then
                lAnz = rsrs!maxi
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        anzeige "normal", "alle Kassenabrechnungen (" & lAnz & "/" & lDiff & ")", Label1(20)
        
        Command5(23).Enabled = True
    Else
        Command5(23).Enabled = False
    End If
    
    'Kassenbons
    If NewTableSuchenDBKombi("KASSBOND", GDPdU_DB) Then
    
        CheckIndex "KASSBOND", "DATUM", "", GDPdU_DB
        'min dat
        Set rsrs = GDPdU_DB.OpenRecordset("select min(DATUM) as ZDatum from KASSBOND")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!ZDatum) Then
                dateStand = rsrs!ZDatum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(37)
        
        'max dat
        Set rsrs = GDPdU_DB.OpenRecordset("select max(DATUM) as ZDatum from KASSBOND")
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!ZDatum) Then
                dateStand = rsrs!ZDatum
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        anzeige "normal", Format(dateStand, "DD.MM.YY"), Label1(36)
        
'        Command5(19).Enabled = True
    Else
'        Command5(19).Enabled = False
    End If
    
    GDPdU_DB.Close
    
    
    
    'alle verfügbaren Bestandshistorien
    fuelle_BestDat
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Aufbereitung_derZahlen"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuelle_BestDat()
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim ctemp       As String
    Dim GDPdU_DB    As Database
    Dim cPfad       As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    cboBestHist.Visible = False
    cboBestHist.Clear
    
    Dim lMAXDatumUebersicht     As Long
    Dim lMAXDatumGLAGER         As Long
    Dim rsDat                   As DAO.Recordset
    
    If NewTableSuchenDBKombi("GLAGER_GDPdU", GDPdU_DB) Then
    
        CheckIndex "GLAGER_GDPdU", "DATUM", "", GDPdU_DB
        CheckIndex "GLAGER_GDPdU", "BESTAND", "", GDPdU_DB
    
        If NewTableSuchenDBKombi("GLAGER_UEBERSICHT", GDPdU_DB) = False Then
            'dann erstelle eine
            cSQL = "select distinct(datum) as disdatum ,sum(bestand) as mBestand into GLAGER_UEBERSICHT from GLAGER_GDPdU group by datum "
            GDPdU_DB.Execute cSQL, dbFailOnError
        Else
            'füge neue Sätze an
            lMAXDatumUebersicht = 0
            lMAXDatumGLAGER = 0
    
            cSQL = "Select Max(disdatum) as Maxdat from GLAGER_UEBERSICHT"
            Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                lMAXDatumUebersicht = rsrs!Maxdat
            End If
            rsrs.Close: Set rsrs = Nothing
            
            cSQL = "Select Max(datum) as Maxdat from GLAGER_GDPdU"
            Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
                lMAXDatumGLAGER = rsrs!Maxdat
            End If
            rsrs.Close: Set rsrs = Nothing
    
            If lMAXDatumGLAGER > lMAXDatumUebersicht Then
                'dann gibt es etwas anzufügen
                cSQL = "Insert into GLAGER_UEBERSICHT select distinct(datum) as disdatum ,sum(bestand) as mBestand "
                cSQL = cSQL & " from GLAGER_GDPdU where datum > " & lMAXDatumUebersicht & " group by datum "
                GDPdU_DB.Execute cSQL, dbFailOnError
                
            End If
        End If
    
        anzeige "", "verfügbare Daten werden ermittelt...", Label1(15)
        
        cSQL = "select  disdatum , mBestand from GLAGER_UEBERSICHT order by disdatum desc"
        Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                If Not IsNull(rsrs!disdatum) Then
                    ctemp = Format(rsrs!disdatum, "DD.MM.YY")
                Else
                    ctemp = ""
                End If
                
                If ctemp <> "" Then
                    If Not IsNull(rsrs!mBestand) Then
                        ctemp = ctemp & " (" & rsrs!mBestand & ")"
                    End If
                
                    If cboBestHist.Text = "" Then
                        cboBestHist.Text = ctemp
                    End If
                    
                    cboBestHist.AddItem ctemp
                End If
                
                rsrs.MoveNext
            Loop
        End If
        rsrs.Close: Set rsrs = Nothing
        Command5(24).Enabled = True
    Else
        Command5(24).Enabled = False
    End If
    
    cboBestHist.Visible = True
    
    GDPdU_DB.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuelle_BestDat"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ExportCSV(sdat As String)
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
    Dim cDatFormat As String
    
    If Text1(7).Text <> "" Then
        cDatFormat = Text1(7).Text
    Else
        cDatFormat = "DDMM"
    End If

   
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sSQL = " Select "
'    sSQL = sSQL & " FILIALE  "
'    sSQL = sSQL & ", FILBEZ  "
    sSQL = sSQL & " ZEITRAUMVON "
    sSQL = sSQL & ", ZEITRAUMBIS "
    sSQL = sSQL & ", KOST  "
    sSQL = sSQL & ", FILKONTO  "
    sSQL = sSQL & ", KONTO  "
    sSQL = sSQL & ", KONTOBEZ  "
    sSQL = sSQL & ", BETRAG "
    sSQL = sSQL & ", GRUND "
    sSQL = sSQL & " from DATEVEXPORT "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        sAusgabedatname = "DATEV" & sdat & ".csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = "ZEITRAUMVON;ZEITRAUMBIS;KOST;FILKONTO;KONTO;KONTOBEZ;BETRAG;GRUND" & Chr$(13) & Chr$(10)

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 7
                If Not IsNull(rsrs.Fields(i)) Then

                    If i > -1 Then
                        If i = 0 Then
                            cSatz = cSatz & "" & Format(rsrs.Fields(i), cDatFormat)
                        ElseIf i = 1 Then
                            cSatz = cSatz & ";" & Format(rsrs.Fields(i), cDatFormat)
                        ElseIf i = 7 Then
                        
                            If rsrs.Fields(i) = 0 Then
                                cSatz = cSatz & ";"
                            Else
                                cSatz = cSatz & ";" & Format(rsrs.Fields(i), "######0.00")
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
    
    If Datendrin("DATEVEXPORT", gdBase) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(4)
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
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ExportCSV_Bauch(sdat As String)
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
    Dim cDatFormat As String
    
    If Text1(7).Text <> "" Then
        cDatFormat = Text1(7).Text
    Else
        cDatFormat = "DDMM"
    End If

   
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
    cPfad1 = gcDBPfad      'dbpfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    
    
    
    
    
'        Case Is = "DATEVEXPORT"
'            cSQL = "Create Table DATEVEXPORT"
'            cSQL = cSQL & " ( "
'            cSQL = cSQL & " ZEITRAUMVON DATETIME"
'            cSQL = cSQL & ", ZEITRAUMBIS DATETIME"
'            cSQL = cSQL & ", KOST varchar(10) "
'            cSQL = cSQL & ", FILKONTO int "
'            cSQL = cSQL & ", KONTO int "
'            cSQL = cSQL & ", KONTOBEZ varchar(35) "
'            cSQL = cSQL & ", BETRAG real"
'            cSQL = cSQL & ", GRUND varchar(50) "
'            cSQL = cSQL & " ) "
'
'        Case Is = "DATEVEXPORT_BAUCH"
'            cSQL = "Create Table DATEVEXPORT_BAUCH"
'            cSQL = cSQL & " ( "
'            cSQL = cSQL & " BETRAG real"
'            cSQL = cSQL & ", BUSCHLUESSEL int "
'            cSQL = cSQL & ", KASSENKONTO int "
'            cSQL = cSQL & ", RENU varchar(36) "
'            cSQL = cSQL & ", BELEGDATUM DATETIME"
'            cSQL = cSQL & ", GEGENKONTO int "
'            cSQL = cSQL & ", BUTEXT varchar(60) "
'            cSQL = cSQL & ", UST int "
'            cSQL = cSQL & ", FESTSCHREIBUNG int "
'            cSQL = cSQL & " ) "
'
    
    
    
    
    loeschNEW "DATEVEXPORT_BAUCH", gdBase
    CreateTableT2 "DATEVEXPORT_BAUCH", gdBase
    
    sSQL = "Insert into DATEVEXPORT_BAUCH select "
    sSQL = sSQL & " BETRAG "
    sSQL = sSQL & ", 0 as BUSCHLUESSEL  "
    sSQL = sSQL & ", Konto as KASSENKONTO  "
    sSQL = sSQL & ", '' as RENU  "
    sSQL = sSQL & ", ZEITRAUMVON as BELEGDATUM "
    sSQL = sSQL & ", 1600 as GEGENKONTO "
    sSQL = sSQL & ", KONTOBEZ as BUTEXT  "
    sSQL = sSQL & ", 19 as UST  "
    sSQL = sSQL & ", 0 as FESTSCHREIBUNG "
    sSQL = sSQL & " from DATEVEXPORT "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DATEVEXPORT_BAUCH set BETRAG =0 where Betrag is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DATEVEXPORT_BAUCH set KASSENKONTO =0 where KASSENKONTO is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DATEVEXPORT_BAUCH set BUTEXT ='' where BUTEXT is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DATEVEXPORT_BAUCH set UST =19 where BUTEXT ='Umsatz volle MwSt'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DATEVEXPORT_BAUCH set UST =7 where BUTEXT ='Umsatz erm MwSt'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update DATEVEXPORT_BAUCH set UST =0 where BUTEXT ='Umsatz ohne MwSt'"
    gdBase.Execute sSQL, dbFailOnError
    
   
    
    sSQL = " Select "
    sSQL = sSQL & " BETRAG "
    sSQL = sSQL & ", BUSCHLUESSEL  "
    sSQL = sSQL & ", KASSENKONTO  "
    sSQL = sSQL & ", RENU  "
    sSQL = sSQL & ", BELEGDATUM "
    sSQL = sSQL & ", GEGENKONTO "
    sSQL = sSQL & ", BUTEXT  "
    sSQL = sSQL & ", UST  "
    sSQL = sSQL & ", FESTSCHREIBUNG  "
    sSQL = sSQL & " from DATEVEXPORT_BAUCH "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        sAusgabedatname = "DATEV" & sdat & ".csv"

        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If

        cdatei = cPfad1 & "BOX\" & sAusgabedatname
        cPfad = cPfad1 & "BOX"
        
        Kill cdatei
        
        iFileNr = FreeFile
        Open cdatei For Binary As #iFileNr
        
        cSatz = "Umsatz (+/-);BU-Schlüssel;Kassenkonto (Konto);Rechnungsnummer (Belegfeld 1);Belegdatum;Sach-/Personalkonto (Gegenkonto: ohne BU-Schlüssel);Buchungstext;USt in % (Inland);Festschreibung" & Chr$(13) & Chr$(10)
'        cSatz = "ZEITRAUMVON;ZEITRAUMBIS;KOST;FILKONTO;KONTO;KONTOBEZ;BETRAG;GRUND" & Chr$(13) & Chr$(10)

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            cSatz = ""
            For i = 0 To 8
            
                If i = 0 Then
                    If rsrs.Fields(i) = 0 Then
                        cSatz = cSatz & ""
                    Else
                        cSatz = cSatz & "" & Format(rsrs.Fields(i), "######0.00")
                    End If
                ElseIf i = 4 Then
                    cSatz = cSatz & ";" & Format(rsrs.Fields(i), "DDMMYYYY")
                Else
                    cSatz = cSatz & ";" & rsrs.Fields(i)
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
    
    If Datendrin("DATEVEXPORT_BAUCH", gdBase) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(4)
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
        Fehler.gsFunktion = "ExportCSV_Bauch"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub GDPdU_ExportCSV_Kassjour(sVondat As String, sBisdat As String)
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
    
    Dim GDPdU_DB        As Database
    
    Dim cDatum          As String
    Dim cArtNr          As String
    Dim czeit           As String
    Dim cBez            As String
    Dim cMenge          As String
    Dim cPreis          As String
    Dim cMWST           As String
    Dim cUms_ok         As String
    Dim cBELEGNR        As String
    Dim cKundnr         As String
    Dim cKasnum         As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(15)
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    sAusgabedatname = "Verkäufe_" & sVondat & "_" & sBisdat & ".csv"
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "GDPdU\" & sAusgabedatname
    cPfad = cPfad1 & "GDPdU"
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cSatz = "Verkaufsdatum;Verkaufszeit;Kassennummer;Artikelnummer;Artikelbezeichnung;Verkaufsmenge;Kassenverkaufspreis(summiert);"
    cSatz = cSatz & "Kundennummer;MwSt;umsatzsteuerpflichtig;Kassenbonnr" & Chr$(13) & Chr$(10)
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "temp_Kassjour", GDPdU_DB
    
    sSQL = "select Adate,Azeit,Artnr,Bezeich,Menge,Preis,Kundnr,MWST,UMS_OK,BELEGNR,kasnum "
    sSQL = sSQL & " into temp_Kassjour from KASSJOUR "
    sSQL = sSQL & " where ADATE >= " & CLng(DateValue(sVondat))
    sSQL = sSQL & " and ADATE <= " & CLng(DateValue(sBisdat))
    sSQL = sSQL & " order by adate, azeit"
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    sSQL = "select Adate,Azeit,Artnr,Bezeich,Menge,Preis,Kundnr,MWST,UMS_OK,BELEGNR,kasnum "
    sSQL = sSQL & " from temp_Kassjour order by adate, azeit"

    Set rsrs = GDPdU_DB.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cDatum = ""
            czeit = ""
            cArtNr = ""
            cBez = ""
            cMenge = ""
            cPreis = ""
            cMWST = ""
            cUms_ok = ""
            cBELEGNR = ""
            cKundnr = ""
            
            If Not IsNull(rsrs!ADATE) Then
                cDatum = rsrs!ADATE
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                czeit = rsrs!AZEIT
            End If
            
            If Not IsNull(rsrs!kasnum) Then
                cKasnum = rsrs!kasnum
            End If
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBez = rsrs!BEZEICH
            End If
            
            If Not IsNull(rsrs!Menge) Then
                cMenge = rsrs!Menge
            End If
            
            If Not IsNull(rsrs!Preis) Then
                cPreis = rsrs!Preis
            End If
            
            If Not IsNull(rsrs!Kundnr) Then
                cKundnr = rsrs!Kundnr
            End If
            
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            End If
            
            If Not IsNull(rsrs!UMS_OK) Then
                cUms_ok = rsrs!UMS_OK
            End If
            
            If Not IsNull(rsrs!BELEGNR) Then
                cBELEGNR = rsrs!BELEGNR
            End If
            
            cSatz = ""
            
            cSatz = cDatum & ";" & czeit & ";" & cKasnum & ";" & cArtNr & ";" & cBez & ";" & cMenge & ";" & cPreis & ";" & cKundnr & ";"
            cSatz = cSatz & cMWST & ";" & cUms_ok & ";" & cBELEGNR
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            
            rsrs.MoveNext
        Loop
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Beschreibung_GDPdU_ExportCSV_Kassjour
    
    If Datendrin("temp_Kassjour", GDPdU_DB) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "GDPdU) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(15)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(15)
    End If
    
    GDPdU_DB.Close
    
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "GDPdU_ExportCSV_Kassjour"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Datev_ExportCSV_Kassjour(sVondat As String, sBisdat As String, sKtnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
'    Dim cPfad           As String
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim iRet            As Integer
    Dim rsrs            As Recordset
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    
'    Dim GDPdU_DB        As Database
    
    Dim cDatum          As String
    Dim cArtNr          As String
    Dim czeit           As String
    Dim cBez            As String
    Dim cMenge          As String
    Dim cPreis          As String
    Dim cMWST           As String
    Dim cUms_ok         As String
    Dim cBELEGNR        As String
    Dim cKundnr         As String
    Dim cKasnum         As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(4)
    
'    cPfad = gcDBPfad
'    If Right$(cPfad, 1) <> "\" Then
'        cPfad = cPfad & "\"
'    End If
'
'    cPfad = cPfad & "GDPdU\GDPdU.MDB"
'
'    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    sAusgabedatname = "Verkäufe_" & sVondat & "_" & sBisdat & ".csv"
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "BOX\" & sAusgabedatname
'    cPfad = cPfad1 & "BOX"
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cSatz = "Verkaufsdatum;Verkaufszeit;Kassennummer;Artikelnummer;Artikelbezeichnung;Verkaufsmenge;Kassenverkaufspreis(summiert);"
    cSatz = cSatz & "Kundennummer;MwSt;umsatzsteuerpflichtig;Kassenbonnr;Konto" & Chr$(13) & Chr$(10)
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "temp_Kassjour", gdBase
    
    sSQL = "select Adate,Azeit,Artnr,Bezeich,Menge,Preis,Kundnr,MWST,UMS_OK,BELEGNR,kasnum "
    sSQL = sSQL & " into temp_Kassjour from KASSJOUR "
    sSQL = sSQL & " where ADATE >= " & CLng(DateValue(sVondat))
    sSQL = sSQL & " and ADATE <= " & CLng(DateValue(sBisdat))
    sSQL = sSQL & " order by adate, azeit"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select Adate,Azeit,Artnr,Bezeich,Menge,Preis,Kundnr,MWST,UMS_OK,BELEGNR,kasnum "
    sSQL = sSQL & " from temp_Kassjour order by adate, azeit"

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cDatum = ""
            czeit = ""
            cArtNr = ""
            cBez = ""
            cMenge = ""
            cPreis = ""
            cMWST = ""
            cUms_ok = ""
            cBELEGNR = ""
            cKundnr = ""
            
            If Not IsNull(rsrs!ADATE) Then
                cDatum = rsrs!ADATE
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                czeit = rsrs!AZEIT
            End If
            
            If Not IsNull(rsrs!kasnum) Then
                cKasnum = rsrs!kasnum
            End If
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBez = rsrs!BEZEICH
            End If
            
            If Not IsNull(rsrs!Menge) Then
                cMenge = rsrs!Menge
            End If
            
            If Not IsNull(rsrs!Preis) Then
                cPreis = rsrs!Preis
            End If
            
            If Not IsNull(rsrs!Kundnr) Then
                cKundnr = rsrs!Kundnr
            End If
            
            If Not IsNull(rsrs!MWST) Then
                cMWST = rsrs!MWST
            End If
            
            If Not IsNull(rsrs!UMS_OK) Then
                cUms_ok = rsrs!UMS_OK
            End If
            
            If Not IsNull(rsrs!BELEGNR) Then
                cBELEGNR = rsrs!BELEGNR
            End If
            
            cSatz = ""
            
            cSatz = cDatum & ";" & czeit & ";" & cKasnum & ";" & cArtNr & ";" & cBez & ";" & cMenge & ";" & cPreis & ";" & cKundnr & ";"
            cSatz = cSatz & cMWST & ";" & cUms_ok & ";" & cBELEGNR & ";" & sKtnr
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            
            rsrs.MoveNext
        Loop
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
'    Beschreibung_GDPdU_ExportCSV_Kassjour
    
    If Datendrin("temp_Kassjour", gdBase) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "BOX) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(4)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(4)
    End If
    
    'GDPdU_DB.Close
    
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Datev_ExportCSV_Kassjour"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Beschreibung_GDPdU_ExportCSV_Kassjour()
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Beschreibung wird erstellt...", Label1(15)
    

    sAusgabedatname = "Beschreibung_Verkäufe_xxx_xxx.txt"
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "GDPdU\" & sAusgabedatname
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cSatz = "Beschreibung der Datei: Verkäufe_xxx_xxx.csv" & Chr$(13) & Chr$(10)
    cSatz = cSatz & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Die Verkäufe_xxx_xxx.csv wird als ASCII - Textdatei und dem Feldtrennzeichen';' Semikolon ausgegeben." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Feldnamen/Erläuterung" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Verkaufsdatum                  Tag des Verkaufsvorganges" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Verkaufszeit                   Uhrzeit des Verkaufsvorganges" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Kassennummer                   Nummer der Kasse" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Artikelnummer                  eindeutige Winkiss Artikelnummer" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Artikelbezeichnung             " & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Verkaufsmenge                  werden 2 Stück a´ 3,00 € verkauft, so steht unter Verkaufsmenge: 2 und unter " & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Kassenverkaufspreis(summiert)  Kassenverkaufspreis 6,00 €" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Kundennummer                   Verkauf mit Kundenbindung enthält eine Kundennummer" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "MwSt                           volle MwSt = 'V' und ermäßigte MwSt = 'E'" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "umsatzsteuerpflichtig          in der Regel = 'J' Ausnahme z.B.: bei Gutscheinverkäufen = 'N'" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Kassenbonnr                    beginnt immer nach einem Kassenabschluss mit 1000" & Chr$(13) & Chr$(10)
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    Close iFileNr
    
    
    
    anzeige "normal", "", Label1(15)
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Beschreibung_GDPdU_ExportCSV_Kassjour"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub GDPdU_ExportCSV_KVKPR1PROT(sVondat As String, sBisdat As String)
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
    
    Dim GDPdU_DB        As Database
    
    Dim cDatum          As String
    Dim cArtNr          As String
    Dim czeit           As String
    Dim cKVKPR          As String
    Dim cBediener       As String
    Dim cAENART         As String
    
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(15)
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    sAusgabedatname = "Preisänderungen_" & sVondat & "_" & sBisdat & ".csv"
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "GDPdU\" & sAusgabedatname
    cPfad = cPfad1 & "GDPdU"
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cSatz = "Änderungsdatum;Änderungszeit;Artikelnummer;neuer Kassenverkaufspreis;"
    cSatz = cSatz & "Bediener;Änderungsart" & Chr$(13) & Chr$(10)
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "temp_KVKPR1PROT", GDPdU_DB
    
    sSQL = " Select "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", KVKPR1  "
    sSQL = sSQL & ", BEDIENER "
    sSQL = sSQL & ", SYNSTATUS  "
    sSQL = sSQL & ", AENART "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & ", LASTDATE "
    sSQL = sSQL & ", LASTTIME  "
    sSQL = sSQL & ", SENDOK  "
    sSQL = sSQL & " into temp_KVKPR1PROT from KVKPR1PROT "
    sSQL = sSQL & " where LASTDATE >= " & CLng(DateValue(sVondat))
    sSQL = sSQL & " and LASTDATE <= " & CLng(DateValue(sBisdat))
    sSQL = sSQL & " order by LASTDATE, LASTTIME"
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    sSQL = " Select "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", KVKPR1  "
    sSQL = sSQL & ", BEDIENER "
    sSQL = sSQL & ", SYNSTATUS  "
    sSQL = sSQL & ", AENART "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & ", LASTDATE "
    sSQL = sSQL & ", LASTTIME  "
    sSQL = sSQL & ", SENDOK  "
    sSQL = sSQL & " from temp_KVKPR1PROT  "
    sSQL = sSQL & " order by LASTDATE, LASTTIME"
    
    Set rsrs = GDPdU_DB.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cDatum = ""
            czeit = ""
            cArtNr = ""
            cKVKPR = ""
            cBediener = ""
            cAENART = ""
            
            If Not IsNull(rsrs!LASTDATE) Then
                cDatum = rsrs!LASTDATE
            End If
            
            If Not IsNull(rsrs!LASTTIME) Then
                czeit = rsrs!LASTTIME
            End If
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!KVKPR1) Then
                cKVKPR = rsrs!KVKPR1
            End If
            
            If Not IsNull(rsrs!BEDIENER) Then
                cBediener = rsrs!BEDIENER
            End If
            
            If Not IsNull(rsrs!AENART) Then
                cAENART = rsrs!AENART
            End If

            cSatz = ""
            
            cSatz = cDatum & ";" & czeit & ";" & cArtNr & ";" & cKVKPR & ";" & cBediener & ";" & cAENART
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            
            rsrs.MoveNext
        Loop
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Beschreibung_GDPdU_ExportCSV_KVKPR1PROT
    
    
    If Datendrin("temp_KVKPR1PROT", GDPdU_DB) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "GDPdU) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(15)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(15)
    End If
    
    GDPdU_DB.Close
    
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "GDPdU_ExportCSV_KVKPR1PROT"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Beschreibung_GDPdU_ExportCSV_KVKPR1PROT()
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Beschreibung wird erstellt...", Label1(15)
    

    sAusgabedatname = "Beschreibung_Preisänderungen_xxx_xxx.txt"
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "GDPdU\" & sAusgabedatname
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    
    
    cSatz = "Beschreibung der Datei: Preisänderungen_xxx_xxx.csv" & Chr$(13) & Chr$(10)
    cSatz = cSatz & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Die Preisänderungen_xxx_xxx.csv wird als ASCII - Textdatei und dem Feldtrennzeichen';' Semikolon ausgegeben." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Feldnamen/Erläuterung" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Änderungsdatum                 Tag der Preisänderung" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Änderungszeit                  Uhrzeit der Preisänderung" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Artikelnummer                  eindeutige Winkiss Artikelnummer" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "neuer Kassenverkaufspreis      geänderter Kassenverkaufspreis" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Bediener                       angemeldeter Bediener" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Änderungsart                   Programmteil in der die Änderung vorgenommen wurde" & Chr$(13) & Chr$(10)
    
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    Close iFileNr
    
    
    
    anzeige "normal", "", Label1(15)
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Beschreibung_GDPdU_ExportCSV_KVKPR1PROT"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub GDPdU_ExportCSV_BESTPROT(sVondat As String, sBisdat As String)
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
    
    Dim GDPdU_DB        As Database
    
    Dim cDatum          As String
    Dim cArtNr          As String
    Dim czeit           As String
    Dim cBediener       As String
    Dim cAENART         As String
    Dim cUMENGE         As String
    Dim cNEWBEST        As String
    Dim cOLDBEST        As String
    Dim cAENGRUND       As String
            
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", Label1(15)
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    sAusgabedatname = "Bestandsänderungen_" & sVondat & "_" & sBisdat & ".csv"
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "GDPdU\" & sAusgabedatname
    cPfad = cPfad1 & "GDPdU"
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cSatz = "Änderungsdatum;Änderungszeit;Artikelnummer;alter Bestand;Bewegungsmenge;neuer Bestand;"
    cSatz = cSatz & "Bediener;Änderungsart;Änderungsgrund" & Chr$(13) & Chr$(10)
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    loeschNEW "temp_BESTPROT", GDPdU_DB
    
    sSQL = " Select "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", UMENGE "
    sSQL = sSQL & ", NEWBEST "
    sSQL = sSQL & ", OLDBEST "
    sSQL = sSQL & ", BEDIENER "
    sSQL = sSQL & ", SYNSTATUS "
    sSQL = sSQL & ", AENART "
    sSQL = sSQL & ", AENGRUND "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & ", LASTDATE "
    sSQL = sSQL & ", LASTTIME "
    sSQL = sSQL & ", SENDOK  "
    sSQL = sSQL & " into temp_BESTPROT from BESTPROT "
    sSQL = sSQL & " where LASTDATE >= " & CLng(DateValue(sVondat))
    sSQL = sSQL & " and LASTDATE <= " & CLng(DateValue(sBisdat))
    sSQL = sSQL & " order by LASTDATE, LASTTIME"
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    sSQL = " Select "
    sSQL = sSQL & " ARTNR "
    sSQL = sSQL & ", UMENGE "
    sSQL = sSQL & ", NEWBEST "
    sSQL = sSQL & ", OLDBEST "
    sSQL = sSQL & ", BEDIENER "
    sSQL = sSQL & ", SYNSTATUS "
    sSQL = sSQL & ", AENART "
    sSQL = sSQL & ", AENGRUND "
    sSQL = sSQL & ", FILIALE "
    sSQL = sSQL & ", LASTDATE "
    sSQL = sSQL & ", LASTTIME "
    sSQL = sSQL & ", SENDOK  "
    sSQL = sSQL & " from temp_BESTPROT  "
    sSQL = sSQL & " order by LASTDATE, LASTTIME"
    
    Set rsrs = GDPdU_DB.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cDatum = ""
            czeit = ""
            cArtNr = ""
            cBediener = ""
            cAENART = ""
            cUMENGE = ""
            cNEWBEST = ""
            cOLDBEST = ""
            cAENGRUND = ""

            If Not IsNull(rsrs!LASTDATE) Then
                cDatum = rsrs!LASTDATE
            End If
            
            If Not IsNull(rsrs!LASTTIME) Then
                czeit = rsrs!LASTTIME
            End If
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            End If
            
            If Not IsNull(rsrs!BEDIENER) Then
                cBediener = rsrs!BEDIENER
            End If
            
            If Not IsNull(rsrs!AENART) Then
                cAENART = rsrs!AENART
            End If
            
            If Not IsNull(rsrs!UMENGE) Then
                cUMENGE = rsrs!UMENGE
            End If
            
            If Not IsNull(rsrs!NEWBEST) Then
                cNEWBEST = rsrs!NEWBEST
            End If
            
            If Not IsNull(rsrs!OLDBEST) Then
                cOLDBEST = rsrs!OLDBEST
            End If
            
            If Not IsNull(rsrs!AENGRUND) Then
                cAENGRUND = rsrs!AENGRUND
            End If
            
            cSatz = ""
            
            cSatz = cDatum & ";" & czeit & ";" & cArtNr & ";" & cOLDBEST & ";" & cUMENGE & ";" & cNEWBEST & ";" & cBediener & ";" & cAENART & ";" & cAENGRUND
            cSatz = cSatz & Chr$(13) & Chr$(10)
            
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
            
            rsrs.MoveNext
        Loop
        Close iFileNr
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Beschreibung_GDPdU_ExportCSV_BESTPROT
    
    
    If Datendrin("temp_BESTPROT", GDPdU_DB) Then
        iRet = MsgBox("Möchten Sie diese CSV - Datei als Email verschicken?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbYes Then
            gcBestellEmail.Attachment1 = cdatei
            Screen.MousePointer = 0
            frmWKL129.Show 1
        Else
            MsgBox "Diese Datei ist unter (" & cPfad1 & "GDPdU) mit dem Namen: " & sAusgabedatname & " abgespeichert", vbInformation, "Winkiss Information:"
        End If
        anzeige "normal", "", Label1(15)
    Else
        anzeige "rot", "Keine Daten zum Export vorhanden.", Label1(15)
    End If
    
    GDPdU_DB.Close
    
    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "GDPdU_ExportCSV_BESTPROT"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Beschreibung_GDPdU_ExportCSV_BESTPROT()
    On Error GoTo LOKAL_ERROR
    
    Dim cdatei          As String
    Dim cPfad1          As String
    Dim sAusgabedatname As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim cSatz           As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Beschreibung wird erstellt...", Label1(15)
    
    sAusgabedatname = "Beschreibung_Bestandsänderungen_xxx_xxx.txt"
    cPfad1 = gcDBPfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If

    cdatei = cPfad1 & "GDPdU\" & sAusgabedatname
    
    Kill cdatei
    
    iFileNr = FreeFile
    Open cdatei For Binary As #iFileNr
    
    cSatz = "Beschreibung der Datei: Bestandsänderungen_xxx_xxx.csv" & Chr$(13) & Chr$(10)
    cSatz = cSatz & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Die Bestandsänderungen_xxx_xxx.csv wird als ASCII - Textdatei und dem Feldtrennzeichen';' Semikolon ausgegeben." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Feldnamen/Erläuterung" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    
    
    cSatz = cSatz & "Änderungsdatum                 Tag der Bestandsänderung" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Änderungszeit                  Uhrzeit der Bestandsänderung" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Artikelnummer                  eindeutige Winkiss Artikelnummer" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "alter Bestand                  Bestand vor der Veränderung" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Bewegungsmenge                 Zu- oder Abgang" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "neuer Bestand                  Bestand nach der Veränderung" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Bediener                       angemeldeter Bediener" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Änderungsart                   Programmteil in der die Änderung vorgenommen wurde" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Änderungsgrund                 bedienerbezogene Angabe bei Bestandsminimierung" & Chr$(13) & Chr$(10)
    
    cSatz = cSatz & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Ausnahme: alle Änderungen bewirken eine Bestandsveränderung mit Ausnahme der Änderungsart: 'Rücknahme Kasse'" & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Mit 'Rücknahme Kasse' werden Artikel gekennzeichnet, die bevor ein Kassenvorgang abgeschlossen wird, aus dem Warenkorb entfernt werden." & Chr$(13) & Chr$(10)
    cSatz = cSatz & "Bestandsveränderungen die durch Verkaufsvorgänge herbeigeführt werden, werden hier nicht aufgeführt." & Chr$(13) & Chr$(10)
    
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    Close iFileNr
    
    
    
    anzeige "normal", "", Label1(15)
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Beschreibung_GDPdU_ExportCSV_BESTPROT"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub LoescheKonten()
On Error GoTo LOKAL_ERROR

    Dim bFound      As Boolean
    Dim sSQL        As String
    Dim cKontobez   As String
    Dim lcount      As Long
    
    bFound = False
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
    
    If Not bFound Then
        anzeige "rot", "Bitte markieren Sie eine Zeile", Label1(4)
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    anzeige "Normal", "", Label1(4)
    
    For lcount = 0 To List2.ListCount - 1
        If List2.Selected(lcount) = True Then
            cKontobez = Trim(Right(List2.list(lcount), 36))
            
            sSQL = "Delete from DATEVKONTEN where KONTOBEZ = '" & cKontobez & "'"
            gdBase.Execute sSQL, dbFailOnError
        End If
    Next
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheKonten"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LoescheKontenAGN()
On Error GoTo LOKAL_ERROR

    Dim bFound      As Boolean
    Dim sSQL        As String
    Dim lagn        As Long
    Dim lcount      As Long
    Dim cKontobez   As String
    
    bFound = False
    
    For lcount = 0 To List3.ListCount - 1
        If List3.Selected(lcount) = True Then
            bFound = True
            Exit For
        End If
    Next lcount
    
    If Not bFound Then
        anzeige "rot", "Bitte markieren Sie eine Zeile", Label1(4)
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    anzeige "Normal", "", Label1(4)
    
    For lcount = 0 To List3.ListCount - 1
        If List3.Selected(lcount) = True Then
            lagn = Trim(Left(List3.list(lcount), 5))
            cKontobez = Trim(Right(List3.list(lcount), 36))
            
            sSQL = "Delete from DATEVKONTENAGN where AGN = " & lagn
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Delete from DATEVALLG where KONTOBEZ = '" & cKontobez & "'"
            gdBase.Execute sSQL, dbFailOnError
        End If
    Next
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheKontenAGN"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub LoescheFilKonten()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String

    sSQL = "Delete from KOST "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheFilKonten"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub ZeigeFilKonten()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cLBSatz     As String
    
    List1.Clear
    
    cSQL = "Select * from KOST order by FILIALE "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!FILIALE) Then
                cFeld = rsrs!FILIALE
            Else
                cFeld = "0"
            End If
            cLBSatz = Space(3 - Len(cFeld)) & cFeld & " "
            
            
            If Not IsNull(rsrs!FILBEZ) Then
                cFeld = rsrs!FILBEZ
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            

            If Not IsNull(rsrs!KOST) Then
                cFeld = rsrs!KOST
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(11 - Len(cFeld))
            
            If Not IsNull(rsrs!FilKonto) Then
                cFeld = rsrs!FilKonto
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(4 - Len(cFeld))
            
            List1.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeFilKonten"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeKonten()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cLBSatz     As String
   
    List2.Clear
    
    cSQL = "Select * from DATEVKONTEN "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!Konto) Then
                cFeld = rsrs!Konto
            Else
                cFeld = ""
            End If
            cLBSatz = cFeld & Space(5 - Len(cFeld))
            
            If Not IsNull(rsrs!KONTOBEZ) Then
                cFeld = rsrs!KONTOBEZ
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKonten"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeKontenAGN()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    Dim cLBSatz     As String
   
    List3.Clear
    
    cSQL = "Select * from DATEVKONTENAGN "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!AGN) Then
                cFeld = rsrs!AGN
            Else
                cFeld = ""
            End If
            cLBSatz = cFeld & Space(6 - Len(cFeld))
            
            If Not IsNull(rsrs!KONTOBEZ) Then
                cFeld = rsrs!KONTOBEZ
            Else
                cFeld = ""
            End If
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            
            List3.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKontenAGN"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ZeigeKontenAuszahlung()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim sGrund      As String
    Dim sKontoNr    As String
    Dim sKontoBez   As String
    Dim i           As Integer
    
    For i = 0 To 9
        Label66(i).Caption = ""
        Text6(i).Text = ""
        Text5(i).Text = ""
    Next i
    
    cSQL = "Select  "
    cSQL = cSQL & " AUSZAHLUNGSGRUND "
    cSQL = cSQL & ", KONTO  "
    cSQL = cSQL & ", KONTOBEZ "
    cSQL = cSQL & " from AUSZAHLUNGSGRUND "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
    
            sGrund = ""
            If Not IsNull(rsrs!AUSZAHLUNGSGRUND) Then
                sGrund = Trim(rsrs!AUSZAHLUNGSGRUND)
            End If
            
            sKontoNr = ""
            If Not IsNull(rsrs!Konto) Then
                sKontoNr = Trim(rsrs!Konto)
            End If
            
            sKontoBez = ""
            If Not IsNull(rsrs!KONTOBEZ) Then
                sKontoBez = Trim(rsrs!KONTOBEZ)
            End If
            
            If sGrund <> "" Then
                For i = 0 To 9
                    If Label66(i).Caption = "" Then
                        Label66(i).Caption = sGrund
                        
                        If sKontoNr <> "" And sKontoBez <> "" Then
                            Text6(i).Text = sKontoNr
                            Text5(i).Text = sKontoBez
                        End If
                        
                        
                        
                        Exit For
                    End If
                Next i
            End If
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeKontenAuszahlung"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherAuszahlungsgrund()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim i           As Integer
    Dim sGrund      As String
    Dim sKontoNr    As String
    Dim sKontoBez   As String
    
    For i = 0 To 9
        sGrund = Label66(i).Caption
        sKontoNr = Text6(i).Text
        sKontoBez = Text5(i).Text
        
        
        If sGrund <> "" And sKontoNr <> "" And sKontoBez <> "" Then
        
            
            sSQL = "Update Auszahlungsgrund set KONTO = " & sKontoNr & " "
            sSQL = sSQL & " ,KONTOBEZ = '" & sKontoBez & "' "
            sSQL = sSQL & " where Auszahlungsgrund = '" & sGrund & "' "
            
            gdBase.Execute sSQL, dbFailOnError
        End If
    
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherAuszahlungsgrund"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherFilKonten(cKost As String, lFilKonto As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
   
    sSQL = "Delete from KOST  "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KOST (KOST,FILKONTO) values ( "
    sSQL = sSQL & "  '" & cKost & "' "
    sSQL = sSQL & " , " & lFilKonto & " "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherFilKonten"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherKonten(lKonto As Long, cKontobez As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
   
    sSQL = "Delete from DATEVKONTEN where KONTOBEZ = '" & cKontobez & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DATEVKONTEN (KONTO,KONTOBEZ) values ( "
    sSQL = sSQL & "  " & lKonto & " "
    sSQL = sSQL & " , '" & cKontobez & "' "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherKonten"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SpeicherKontenAGN(lagn As Long, cKontobez As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
   
    sSQL = "Delete from DATEVKONTENAGN where KONTOBEZ = '" & cKontobez & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from DATEVALLG where KONTOBEZ = '" & cKontobez & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DATEVKONTENAGN (AGN,KONTOBEZ) values ( "
    sSQL = sSQL & "  " & lagn & " "
    sSQL = sSQL & " , '" & cKontobez & "' "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('" & cKontobez & "') "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherKontenAGN"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermKonto(cKontobez As String) As Long
    On Error GoTo LOKAL_ERROR
    
    ermKonto = 0
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    sSQL = "Select KONTO from DATEVKONTEN "
    sSQL = sSQL & " where KONTOBEZ = '" & cKontobez & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Konto) Then
            ermKonto = rsrs!Konto
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKonto"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermKOST() As String
    On Error GoTo LOKAL_ERROR
    
    ermKOST = ""
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    sSQL = "Select KOST from KOST "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!KOST) Then
            ermKOST = rsrs!KOST
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermKOST"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermFilKonto() As Long
    On Error GoTo LOKAL_ERROR
    
    ermFilKonto = 0
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    sSQL = "Select FilKonto from KOST  "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!FilKonto) Then
            ermFilKonto = rsrs!FilKonto
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermFilKonto"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub EXPORT(lTag As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim dBetrag         As Double
    Dim dECBetrag       As Double
    Dim dLSBetrag       As Double
    Dim cKostenstelle   As String
    Dim lFilKonto       As Long
    Dim lMaxBenutzerdef As Long
    Dim lKonto          As Long
    Dim i               As Integer
    Dim rsrs            As Recordset
    Dim lagn            As Long
    Dim cKontobez       As String
    
    cKostenstelle = ermKOST()
    lFilKonto = ermFilKonto()
    
    'Umsatz 19%
    dBetrag = ermgesUmsatzMwstAusZumsatz(CStr(lTag), CStr(lTag), "V")
    lKonto = ermKonto("Umsatz volle MwSt")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz volle MwSt", dBetrag
    
    'Umsatz 7%
    dBetrag = ermgesUmsatzMwstAusZumsatz(CStr(lTag), CStr(lTag), "E")
    lKonto = ermKonto("Umsatz erm MwSt")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz erm MwSt", dBetrag
    
    'Umsatz ohne
    dBetrag = ermgesUmsatzMwstAusZumsatz(CStr(lTag), CStr(lTag), "O")
    lKonto = ermKonto("Umsatz ohne MwSt")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz ohne MwSt", dBetrag
    
    'KK Gesamt
    dBetrag = -1 * ermgesKKgesamt(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("Kreditkarten gesamt")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten gesamt", dBetrag
    
    'KK Gesamt V
    dBetrag = -1 * ermgesKKgesamtMwst(CStr(lTag), CStr(lTag), "V")
    lKonto = ermKonto("Kreditkarten ges V")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten ges V", dBetrag
    
    'KK Gesamt E
    dBetrag = -1 * ermgesKKgesamtMwst(CStr(lTag), CStr(lTag), "E")
    lKonto = ermKonto("Kreditkarten ges E")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten ges E", dBetrag
    
    'KK Gesamt O
    dBetrag = -1 * ermgesKKgesamtMwst(CStr(lTag), CStr(lTag), "O")
    lKonto = ermKonto("Kreditkarten ges O")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten ges O", dBetrag
    
    'KK SO
    dBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), "SO")
    lKonto = ermKonto("Kreditkarten Sonstige")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten Sonstige", dBetrag
    
    'KK AE
    dBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), "AE")
    lKonto = ermKonto("Kreditkarten Amex")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten Amex", dBetrag
    
    'KK VI
    dBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), "VI")
    lKonto = ermKonto("Kreditkarten Visa")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten Visa", dBetrag
    
    'KK EU
    dBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), "EU")
    lKonto = ermKonto("Kreditkarten Eurocard")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditkarten Eurocard", dBetrag
    
    'Kreditverkauf
    dBetrag = -1 * ermgesKREDAusZumsatz(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("Kreditverkauf")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kreditverkauf", dBetrag
    
    'Kassensaldo
    dBetrag = ermKassensaldo(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("Kassensaldo")
    If lKonto > 0 Then
        InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kassensaldo", dBetrag
    End If
    
    'Geldtransit zur Bank
    dBetrag = -1 * ermgesABSCHOPF(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("Geldtransit zur Bank")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Geldtransit zur Bank", dBetrag
    
    'Kassendifferenzen
    dBetrag = ermgesKassendiff(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("Kassendifferenzen")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kassendifferenzen", dBetrag
    
    'KK EC
    dECBetrag = -1 * ermgesKK(CStr(lTag), CStr(lTag), "EC")
    
    'Lastschriften
    dLSBetrag = -1 * ermgesLASTZAHLTE(CStr(lTag), CStr(lTag))
    
    dBetrag = dECBetrag + dLSBetrag
    lKonto = ermKonto("Karte EC/ELV")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Karte EC/ELV", dBetrag
    
'    'Umsatz aus Gutscheinlösung (nur der Umsatzanteil aller eingereichten Gutscheine)
'    dBetrag = ermUmsatzAusEingereichtenGutscheinen(CStr(lTag), CStr(lTag))
'    lKonto = ermKonto("Umsatz aus Gutscheinen")
'    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz aus eingereichten Gutscheinen", dBetrag
    
    'Kredit Tilgungen
    dBetrag = ermKreditTilgung(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("Kredit Tilgungen")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Kredit Tilgungen", dBetrag
    
    'VK Restgutscheine
    dBetrag = ermVK_RESTGUTSCH(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("VK Restgutscheine")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "VK Restgutscheine", dBetrag
    
    'Gutscheinlösung (alle eingereichten Gutscheine)
    dBetrag = -1 * ermgesGUTZ(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("Gutscheine")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "komplett eingereichte Gutscheine", dBetrag
    
    'VK Gutschein über Gutschein
    dBetrag = ermGutschausGutsch(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("VK Gutschein über Gutschein")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "VK Gutschein über Gutschein", dBetrag

    'VK Gutschein über Karte
    dBetrag = ermGutschausKarte(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("VK Gutschein über Karte")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "VK Gutschein über Karte", dBetrag
    
    'VK Gutschein über Kredite
    dBetrag = ermGutschausKred(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("VK Gutschein über Kredite")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "VK Gutschein über Kredite", dBetrag
    
    'VK Gutschein über Bar
    dBetrag = ermGutschausBar(CStr(lTag), CStr(lTag))
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("VK Gutschein über Bar")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "VK Gutschein über Bar", dBetrag
    
    'neu
    'VK Gutschein über Scheck
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "GUTSCHSCH")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("VK Gutschein über Scheck")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "VK Gutschein über Scheck", dBetrag
    
    'neu
    'VK Gutschein über Lastschrift
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "GUTSCHLAST")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("VK Gutschein über Lastschrift")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "VK Gutschein über Lastschrift", dBetrag
    
    'nicht umsrel. VK
    dBetrag = ermNichtumsatzRelevant(CStr(lTag), CStr(lTag))
    lKonto = ermKonto("nicht umsrel. VK")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "nicht umsrel. VK", dBetrag
    
    
    
    'neu
    'Umsatz Bar
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "UMS_BAR")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("Umsatz Bar")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Bar", dBetrag
    
    
    
    
    
    'Umsatz Bar V
    dBetrag = -1 * ermgesBARgesamtMwst(CStr(lTag), CStr(lTag), "V")
    lKonto = ermKonto("Umsatz Bar V")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Bar V", dBetrag
    
    'Umsatz Bar E
    dBetrag = -1 * ermgesBARgesamtMwst(CStr(lTag), CStr(lTag), "E")
    lKonto = ermKonto("Umsatz Bar E")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Bar E", dBetrag
    
    'Umsatz Bar O
    dBetrag = -1 * ermgesBARgesamtMwst(CStr(lTag), CStr(lTag), "O")
    lKonto = ermKonto("Umsatz Bar O")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Bar O", dBetrag
    
    
    
    
    
    
    
    'neu
    'Umsatz Scheck
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "UMS_SCHECK")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("Umsatz Scheck")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Scheck", dBetrag
    
    
    
    
    
    
    
    
    'neu
    'Umsatz Kreditkarten
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "UMS_KARTE")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("Umsatz Kreditkarten")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Kreditkarten", dBetrag
    
    'neu
    'Umsatz Gutscheinen
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "ZHLGGUTSCH")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("Umsatz Gutscheinen")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Gutscheinen", dBetrag
    
    
    'neu
    'Umsatz Lastschrift
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "UMS_LAST")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("Umsatz Lastschrift")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Umsatz Lastschrift", dBetrag
    
    
    
    'neu
    'Gutscheinauszahlung
    dBetrag = ermX_ausAFCSTATP(CStr(lTag), CStr(lTag), "AUSZGUTSCH")
    dBetrag = Format(dBetrag, "######0.00")
    lKonto = ermKonto("Gutscheinauszahlung")
    InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, "Gutscheinauszahlung", dBetrag
    
    
 
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    'Ausgaben Auszahlungen
    lKonto = ermKonto("Ausgaben")
    schreibe_alle_EINAUSZAHLUNGen lTag, "AUSZAHLUNG", cKostenstelle, lFilKonto, lKonto
    
    'Einzahlungen
    lKonto = ermKonto("Einzahlungen")
    schreibe_alle_EINAUSZAHLUNGen lTag, "EINZAHLUNG", cKostenstelle, lFilKonto, lKonto

    
    
    
    
    
    'benutzerdefinierte Konten auf AGN
    
    
    
    Set rsrs = gdBase.OpenRecordset("DATEVKONTENAGN")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!KONTOBEZ) Then
                cKontobez = rsrs!KONTOBEZ
            End If
            
            If Not IsNull(rsrs!AGN) Then
                lagn = rsrs!AGN
            End If
            
            dBetrag = ermgesBenuAGN(CStr(lTag), CStr(lTag), lagn)
            lKonto = ermKonto(cKontobez)
            InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, cKontobez, dBetrag
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    'benutzerdefinierte Konten auf Auszahlungsgründe
    
    
    Dim sGrund As String
    
    Set rsrs = gdBase.OpenRecordset("AUSZAHLUNGSGRUND")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            cKontobez = ""
            If Not IsNull(rsrs!KONTOBEZ) Then
                cKontobez = rsrs!KONTOBEZ
            End If
            
            sGrund = ""
            If Not IsNull(rsrs!AUSZAHLUNGSGRUND) Then
                sGrund = rsrs!AUSZAHLUNGSGRUND
            End If
            
            lKonto = 0
            If Not IsNull(rsrs!Konto) Then
                lKonto = Val(rsrs!Konto)
            End If
            
            If lKonto > 0 And cKontobez <> "" Then
                dBetrag = ermgesAuszahlungsgrundDatev(CStr(lTag), CStr(lTag), sGrund)
            
                InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, cKontobez, dBetrag
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    anzeige "normal", "", Label1(4)
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EXPORT"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub schreibe_alle_EINAUSZAHLUNGen(lTag As Long, sArt As String _
, cKostenstelle As String, lFilKonto As Long, lKonto As Long)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim dBetrag     As Double
    Dim sGrund      As String
    
    
    sSQL = "Select * " 'sum(BETRAG) as Maxi
    sSQL = sSQL & " from KAEINAUSF "
    sSQL = sSQL & " where ADATE = " & lTag
    sSQL = sSQL & " and ART = '" & sArt & "'"
    sSQL = sSQL & " and BEZEICH <>  'KB - Korrektur' "
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            dBetrag = 0
            If Not IsNull(rsrs!Betrag) Then
                dBetrag = rsrs!Betrag
            End If
            
            sGrund = ""
            If Not IsNull(rsrs!BEZEICH) Then
                sGrund = rsrs!BEZEICH
            End If
            
            If sArt = "AUSZAHLUNG" Then
                dBetrag = -1 * dBetrag
            End If
            
            InsertDATEVEXPORT Format(lTag, "DD.MM.YY"), cKostenstelle, lFilKonto, lKonto, sArt, dBetrag, sGrund
            
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close
                
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "schreibe_alle_EINAUSZAHLUNGen"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub InsertDATEVEXPORT(cTag As String, cKost As String, lFilKonto As Long _
, lKonto As Long, cKontobez As String, dBetrag As Double, Optional sGrund As String)
On Error GoTo LOKAL_ERROR

    If dBetrag = 0 Then
        Exit Sub
    End If
    
    If lKonto = 0 Then
        Exit Sub
    End If

    Dim sSQL        As String
    
    sSQL = "Insert into DATEVEXPORT ( "
'    sSQL = sSQL & " FILIALE  "
'    sSQL = sSQL & ", FILBEZ  "
    sSQL = sSQL & " ZEITRAUMVON "
    sSQL = sSQL & ", ZEITRAUMBIS "
    sSQL = sSQL & ", KOST  "
    sSQL = sSQL & ", FILKONTO  "
    
    sSQL = sSQL & ", KONTO  "
    sSQL = sSQL & ", KONTOBEZ  "
    sSQL = sSQL & ", BETRAG "
    sSQL = sSQL & ", GRUND "
    sSQL = sSQL & " ) "
    sSQL = sSQL & " values ( "
'    sSQL = sSQL & ifilnr
'    sSQL = sSQL & ", '" & ermFilBez(CLng(ifilnr)) & "'"
    sSQL = sSQL & " '" & cTag & "'"
    sSQL = sSQL & ", '" & cTag & "'"
    sSQL = sSQL & ", '" & cKost & "'"
    sSQL = sSQL & ", " & lFilKonto & "  "
    
    sSQL = sSQL & ", " & lKonto & "  "
    sSQL = sSQL & ", '" & cKontobez & "'"
    sSQL = sSQL & ", '" & dBetrag & "'  "
    sSQL = sSQL & ", '" & sGrund & "'  "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertDATEVEXPORT"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichernE171()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cDatFormat As String
    Dim bo1     As Integer
    
    
    loeschNEW "E171", gdBase
    CreateTableT2 "E171", gdBase
    
    If Trim(Text1(7).Text) = "" Then
        cDatFormat = "DDMM"
    Else
        cDatFormat = Text1(7).Text
    End If
    
    
    If Check1.value = vbChecked Then
        bo1 = 0
    Else
        bo1 = -1
    End If

    sSQL = "Insert into E171 (DATFORMAT,bo1) values "
    sSQL = sSQL & "('" & cDatFormat & "'," & bo1 & ")"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichernE171"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladenE171()
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    
    Text1(7).Text = "DDMM"
    Check1.value = vbUnchecked
    
    If NewTableSuchenDBKombi("E171", gdBase) Then
    
        Set rs = gdBase.OpenRecordset("E171")
        If Not rs.EOF Then
        
            If Not IsNull(rs!DATFORMAT) Then
                Text1(7).Text = rs!DATFORMAT
            Else
                Text1(7).Text = "DDMM"
            End If
            
            
            If rs!bo1 = True Then
                Check1.value = vbUnchecked
            Else
                Check1.value = vbChecked
            End If
        
        
        End If
        rs.Close: Set rs = Nothing
    
    End If

     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladenE171"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim i           As Integer
    
    PositionierenZ174
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    Dim sStandText(36) As String
    
    If NewTableSuchenDBKombi("KOST", gdBase) = False Then
        CreateTableT2 "KOST", gdBase
    End If
    
    If NewTableSuchenDBKombi("DATEVKONTEN", gdBase) = False Then
        CreateTableT2 "DATEVKONTEN", gdBase
    End If
    
    If NewTableSuchenDBKombi("DATEVKONTENAGN", gdBase) = False Then
        CreateTableT2 "DATEVKONTENAGN", gdBase
    End If
    
    If NewTableSuchenDBKombi("AUSZAHLUNGSGRUND", gdBase) = False Then
        CreateTableT2 "AUSZAHLUNGSGRUND", gdBase
    End If
    
    If NewTableSuchenDBKombi("DATEVALLG", gdBase) = False Then
        CreateTableT2 "DATEVALLG", gdBase
    End If
    
    sStandText(0) = "Umsatz volle MwSt"
    sStandText(1) = "Umsatz erm MwSt"
    sStandText(2) = "Umsatz ohne MwSt"
    sStandText(3) = "Kreditkarten Amex"
    sStandText(4) = "Kreditkarten Eurocard"
    sStandText(5) = "Kreditkarten Visa"
    
    sStandText(6) = "Karte EC/ELV"
    sStandText(7) = "Kreditverkauf"
    sStandText(8) = "Geldtransit zur Bank"
    sStandText(9) = "Ausgaben"
    sStandText(10) = "Kassendifferenzen"
    sStandText(11) = "Gutscheine"
    
    sStandText(12) = "nicht umsrel. VK"
    sStandText(13) = "VK Gutschein über Bar"
    sStandText(14) = "VK Gutschein über Kredite"
    sStandText(15) = "VK Gutschein über Karte"
    sStandText(16) = "VK Gutschein über Gutschein"
    
    sStandText(17) = "VK Gutschein über Scheck" 'neu
    sStandText(18) = "VK Gutschein über Lastschrift" 'neu
    
    
    sStandText(19) = "Einzahlungen"
    sStandText(20) = "Kassensaldo"
    sStandText(21) = "VK Restgutscheine"
    sStandText(22) = "Kredit Tilgungen"
    sStandText(23) = "Kreditkarten gesamt"
    
    sStandText(24) = "Umsatz Bar" 'neu
    sStandText(25) = "Umsatz Scheck" 'neu
    sStandText(26) = "Umsatz Kreditkarten" 'neu
    sStandText(27) = "Umsatz aus Gutscheinen" 'neu
    sStandText(28) = "Umsatz aus Lastschrift" 'neu
    
    sStandText(29) = "Gutscheinauszahlung" 'neu
    sStandText(30) = "Kreditkarten Sonstige"

    sStandText(31) = "Kreditkarten ges V" 'neu für Leers , gleichzeitig Unfug
    sStandText(32) = "Kreditkarten ges E"
    sStandText(33) = "Kreditkarten ges O"
    
    sStandText(34) = "Umsatz Bar V" 'neu für Leers , gleichzeitig Unfug
    sStandText(35) = "Umsatz Bar E"
    sStandText(36) = "Umsatz Bar O"
    
    For i = 0 To 36
        sSQL = "Delete from DATEVALLG where KONTOBEZ = '" & sStandText(i) & "'"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into DATEVALLG (KONTOBEZ) values ('" & sStandText(i) & "') "
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    Text1(9).Text = Format(DateValue(Now), "DD.MM.YYYY")
    Text1(3).Text = Format(DateValue(Now), "DD.MM.YYYY")
    
    If NewTableSuchenDBKombi("E171", gdBase) Then
        
   
        If SpalteInTabellegefundenNEW("E171", "bo1", gdBase) = False Then
            SpalteAnfuegenNEW "E171", "bo1", "BIT", gdBase
            
            sSQL = "Update E171 set bo1 = -1 "
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    End If
    voreinstellungladenE171
   
    anzeige "normal", "", Label1(4)
    anzeige "normal", "", Label1(15)
    
    Dim sText As String
    
    sText = "Um an die gewünschten Daten bzw. Informationen für das Finanzamt zu kommen, gehen Sie bitte wie folgt vor:" & vbCrLf
    sText = sText & "Wählen Sie die GDPDU Komplettausgabe." & vbCrLf
    sText = sText & "Nun den gewünschten Zeitraum eingeben und auf EXPORT drücken. Nach kurzer Rechenzeit verrät Ihnen das Programm, wo sich die nun erstellte GDPDU_xxxxxxxx_xxxxxxxx.MDB befindet. Diese kopieren Sie bitte für den Prüfer auf einen USB Stickt. Die Datei ist im MBD Format erstellt." & vbCrLf & vbCrLf

    sText = sText & "Sie können neben der Komplettausgabe die Tabellen: Verkäufe, Preisänderungen und Bestandsänderungen auch einzeln ausgeben. Hierzu gehen Sie wie folgt vor:" & vbCrLf
    sText = sText & "Wählen Sie den Punkt GDPDU. Geben Sie den Zeitraum ein, und klicken jeweils auf CSV. Auch hier teilt Ihnen WINKISS mit, wo sich die erstellten Daten befinden.  Auch diese kopieren Sie bitte auf einen USB Stick. Diese Dateien entsprechen dem CSV Format." & vbCrLf & vbCrLf

    sText = sText & "Um an die Satzbeschreibungen zu kommen, gehen Sie bitte wie folgt vor: Nach der GDPDU Komplettausgabe (s. oben) klicken Sie auf 'Beschreibung'. Sie werden weitergeleitet auf unsere Homepage, von wo aus Sie sich die Satzbeschreibungen ausdrucken oder als PDF Dateien speichern können." & vbCrLf
    sText = sText & "Ebenfalls können Sie hier ein Handbuch von WINKISS herunterladen." & vbCrLf & vbCrLf

    sText = sText & "Ansprechpartner für das Finanzamt sind Herr Heinz, Frau Bohling oder Herr Schröder. Wir geben den Prüfern gerne zu allen Fragen Antworten." & vbCrLf & vbCrLf
    
    
    anzeige "normal", sText, Label1(44)
    
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fülleDATEVALLG(cbox As ComboBox)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    cbox.Clear
    cbox.AddItem "bitte auswählen"
    cbox.Text = "bitte auswählen"

    sSQL = "Select * from DATEVALLG "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!KONTOBEZ) Then
                cbox.AddItem rsrs!KONTOBEZ
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllefil1"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenZ174()
On Error GoTo LOKAL_ERROR

    With Frame1
        .Height = 8655
        .Left = 0
        .Top = 0
        .Width = 11895
    End With
    
    With Frame5
        .Height = 8655
        .Left = 0
        .Top = 0
        .Width = 11895
    End With
    
    With Frame6
        .Height = 8655
        .Left = 0
        .Top = 0
        .Width = 11895
    End With
    
    With Frame8
        .Height = 8655
        .Left = 0
        .Top = 0
        .Width = 11895
    End With

    With Frame4
        .Height = 5535
        .Left = 5040
        .Top = 1560
        .Width = 6615
    End With
    
    With Frame7
        .Height = 5535
        .Left = 5040
        .Top = 1560
        .Width = 6615
    End With
    
    With Frame2
        .Height = 6735
        .Left = 120
        .Top = 1560
        .Width = 11655
    End With
    
    With Frame3
        .Height = 6735
        .Left = 120
        .Top = 1560
        .Width = 11655
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenZ174"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

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
Private Sub drucke_nachträglich_Zbon(lZbonNr As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim cPfad       As String
    Dim GDPdU_DB    As Database
    Dim rsrs        As DAO.Recordset
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Druckvorschau wird erstellt...", Label1(15)
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    loeschNEW "PRINT_ZBON", GDPdU_DB
    CreateTableT2 "PRINT_ZBON", GDPdU_DB
    
    sSQL = "Insert into PRINT_ZBON Select * from TAGKOPF_TEMP where neueANR = " & lZbonNr
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    loeschNEW "PRINT_ZBON", gdBase
    TransferTab GDPdU_DB, gcDBPfad & "\kissdata.mdb", "PRINT_ZBON"
    
    reportbildschirm "", "aWKL171"
    
    anzeige "normal", "", Label1(15)

    Screen.MousePointer = 0
    GDPdU_DB.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "drucke_nachträglich_Zbon"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(16).ForeColor = glS1
    Label1(17).ForeColor = glS1
    Label1(18).ForeColor = glS1
    Label1(22).ForeColor = glS1
    Label1(23).ForeColor = glS1
    Label1(25).ForeColor = glS1
    Label1(26).ForeColor = glS1
    Label1(29).ForeColor = glS1
    Label1(30).ForeColor = glS1
    Label1(36).ForeColor = glS1
    Label1(37).ForeColor = glS1
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame6_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    For i = 0 To 10
        Label4(i).ForeColor = glS1
    Next i
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame8_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Label1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim cNr As String

    Select Case index
        Case 16, 17, 22, 23, 25, 26, 36, 37
            Label1(index).Caption = Format(Datumschreiben11a(3000, 4000), "DD.MM.YY")
        Case 18
            URLGoTo Me.hwnd, "http://www.kisslive.de/downloads/winkiss/Handbuch.pdf"
        Case 30, 29
            cNr = Label1(index).Caption
            
            cNr = SwapStr(cNr, "Nr. ", "")
            
            drucke_nachträglich_Zbon CLng(cNr)
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Select Case index
        Case 16, 17, 18, 22, 23, 25, 26, 29, 30, 36, 37
            Label1(index).ForeColor = glWarn
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Label4_Click(index As Integer)
On Error GoTo LOKAL_ERROR


    Select Case index
        Case 0
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Umsatz"
        Case 1
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Verkauf"
        Case 2
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#EinAus"
        Case 3
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Bediener"
        Case 4
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Gutscheine"
        Case 5
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Kassenbuch"
        Case 6
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Zugang"
        Case 7
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Artikel"
        Case 8
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Lieferanten"
        Case 9
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/winkiss-beitraege/246-gdpdu_schnittstelle#Kunden"
        
            
        Case 10
            URLGoTo Me.hwnd, "http://www.kisslive.de/downloads/winkiss/Handbuch.pdf"
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_Click"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Label4_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    For i = 0 To 10
        Label4(i).ForeColor = glS1
    Next i
    
    Label4(index).ForeColor = glWarn
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub Option2_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    If Option2(0).value = True Then
        Option2(0).ForeColor = glWarn
        Option2(2).ForeColor = glS1
    ElseIf Option2(2).value = True Then
        Option2(2).ForeColor = glWarn
        Option2(0).ForeColor = glS1
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Option2_Click"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case index
        Case 6, 2, 1, 0, 11
            cValid = "1234567890" & Chr$(8)
        Case 7
            cValid = "DMY." & Chr$(8)
    End Select

    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text5_KeyPress(index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String

    Select Case index
        Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46)  '& - .
            cValid = cValid & "+äÄÜüÖöß/:\%()"
    End Select

    cZeichen = Chr$(KeyAscii)
    cZeichen = UCase$(cZeichen)

    If InStr(cValid, cZeichen) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(cZeichen)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

