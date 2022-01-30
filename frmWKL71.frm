VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Begin VB.Form frmWKL71 
   Caption         =   "Artikelliste aus MDE"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmWKL71.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   2655
      Left            =   6360
      TabIndex        =   31
      Top             =   600
      Width           =   5535
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   1
         Left            =   9480
         TabIndex        =   34
         Top             =   6120
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
         Caption         =   "weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Artikelliste mit MDE - Gerät"
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
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Value           =   -1  'True
         Width           =   6975
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Artikelliste mit Scanner"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   6975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "Aktualisieren der Artikel - Stammdaten von MDE-Geräten"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   8520
         MouseIcon       =   "frmWKL71.frx":0442
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   47
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Hilfethemen:"
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
         Index           =   32
         Left            =   8520
         TabIndex        =   46
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label15 
         Caption         =   "Wie möchten Sie vorgehen?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   5775
      Left            =   -600
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   9600
         TabIndex        =   43
         Top             =   3600
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "Schnitt EK"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   45
            Top             =   600
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Listen EK"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   44
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Excelliste füllen"
         Height          =   375
         Left            =   10200
         TabIndex        =   40
         Top             =   360
         Width           =   1455
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   5
         Left            =   9600
         TabIndex        =   39
         Top             =   4920
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   4
         Left            =   9600
         TabIndex        =   38
         Top             =   5520
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   3
         Left            =   9600
         TabIndex        =   37
         Top             =   6120
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
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   1440
         MaxLength       =   13
         TabIndex        =   21
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   120
         MaxLength       =   4
         TabIndex        =   20
         Text            =   "1"
         Top             =   1200
         Width           =   855
      End
      Begin sevCommand3.Command Command2 
         Height          =   525
         Index           =   11
         Left            =   4920
         TabIndex        =   19
         Top             =   1200
         Width           =   2175
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
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   6975
      End
      Begin VB.CheckBox Check9 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "nach dem Scannen Menge = 1"
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   3135
      End
      Begin sevCommand3.Command Command8 
         Height          =   230
         Left            =   1080
         TabIndex        =   41
         Top             =   1500
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   397
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
      Begin sevCommand3.Command Command7 
         Height          =   230
         Left            =   1080
         TabIndex        =   42
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   397
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
      Begin VB.Label Label7 
         Caption         =   "Scanner"
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
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "EAN"
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
         Index           =   5
         Left            =   1560
         TabIndex        =   29
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "Menge"
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
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Ihr Vorgang enthält:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "verschiedene Artikel:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "mit einem Gesamtbestand:"
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
         TabIndex        =   25
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "0"
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
         Index           =   10
         Left            =   4320
         TabIndex        =   24
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "0"
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
         Index           =   11
         Left            =   4320
         TabIndex        =   23
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   7320
         X2              =   7320
         Y1              =   6840
         Y2              =   240
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   8040
         TabIndex        =   22
         Top             =   240
         Width           =   3735
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   0
      Left            =   9600
      TabIndex        =   0
      Top             =   7800
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
      Caption         =   "Schließen"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame10"
      Height          =   6615
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Frame Frame2 
         Caption         =   "Pfad prüfen"
         Height          =   1455
         Left            =   7200
         TabIndex        =   48
         Top             =   2400
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton Command3 
            Caption         =   "Nein"
            Height          =   435
            Left            =   2400
            TabIndex        =   52
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Ja"
            Height          =   435
            Left            =   1440
            TabIndex        =   51
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TextConvPfad 
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Label2 
            Caption         =   "ist der Pfad richtig ?"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   3015
         End
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   2
         Left            =   9600
         TabIndex        =   36
         Top             =   4320
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
      Begin sevCommand3.Command Command2 
         Height          =   315
         Index           =   21
         Left            =   5640
         TabIndex        =   15
         Top             =   6240
         Width           =   1455
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
         Caption         =   "Leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   14
         Top             =   4920
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
         Caption         =   "Betanken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   6975
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   22
         Left            =   9600
         TabIndex        =   7
         Top             =   6120
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
         Caption         =   "Drucken"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command2 
         Height          =   495
         Index           =   23
         Left            =   9600
         TabIndex        =   6
         Top             =   5520
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
         Caption         =   "Einlesen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox txtStatus 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picprogress 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   10635
         TabIndex        =   4
         Top             =   480
         Width           =   10695
      End
      Begin VB.Image Image2 
         Height          =   915
         Left            =   10920
         MouseIcon       =   "frmWKL71.frx":074C
         MousePointer    =   99  'Benutzerdefiniert
         Picture         =   "frmWKL71.frx":0A56
         ToolTipText     =   "Klicken Sie hier, wenn Sie Daten aus dem MDE - Gerät einlesen möchten"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "0"
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
         Index           =   21
         Left            =   3840
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "0"
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
         Index           =   22
         Left            =   3840
         TabIndex        =   12
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "mit einem Gesamtbestand:"
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
         Index           =   23
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Rechts
         Caption         =   "verschiedene Artikel:"
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
         Index           =   24
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label7 
         Caption         =   "Ihr Vorgang enthält:"
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
         Index           =   25
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   5055
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Width           =   9375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Artikelliste aus MDE / Scanner"
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
      TabIndex        =   1
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frmWKL71"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mdeErr As Boolean
Dim gArt As String

Private Sub Command1_Click()

gsConverterPfad = TextConvPfad.Text
Frame2.Visible = False
Screen.MousePointer = 0
If gsMDEGERAET = "FORCOM" Then
    frmWKL93.Show 1
 ElseIf gsMDEGERAET = "REWEMDE" Then
    frmWKL183.Show 1
ElseIf gsMDEGERAET = "CIPHERLAB" Then
    frmWKL183.Show 1
End If
            
End Sub

Private Sub Command2_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim iRet        As Integer
    Dim sSQL        As String
    Dim rsrs        As DAO.Recordset
    Dim dMINLEKPR   As Double
    Dim sArt        As String
    
    Screen.MousePointer = 11
    
    Select Case index
        Case 0
            gsConverterPfad = TextConvPfad.Text
            Frame2.Visible = True
            
        Case 1
            If Option2(0).value = True Then
                Frame10.Visible = True
                Frame6.Visible = False
            ElseIf Option2(1).value = True Then
                Frame9.Visible = True
                Frame6.Visible = False
                Text3(0).SetFocus
            End If
        Case 2
            Frame10.Visible = False
            Frame6.Visible = True
        Case 3
            Frame9.Visible = False
            Frame6.Visible = True
        Case 4 'Druckvorschau aus scanner
        
            If Not NewTableSuchenDBKombi(srechnertab & "ATI", gdBase) Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        
            loeschNEW "BWINVI", gdBase
            CreateTableT2 "BWINVI", gdBase
            
            sSQL = "Insert into BWINVI select ARTNR,LINR,BEZEICH,BESTAND,LEKPR,EKPR,SEKWERT,LEKWERT,ean,ean2,ean3 from " & srechnertab & "ATI "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update BWINVI inner join Artikel "
            sSQL = sSQL & "  on BWINVI.ARTNR = Artikel.artnr "
            sSQL = sSQL & " set BWINVI.MWST = Artikel.MWST "
            gdBase.Execute sSQL, dbFailOnError
            
            If Option1(0).value = True Then
            
                'Kleinsten Listen_EK ausser 0.00 updaten
                
                sSQL = "Select a.Artnr, Min(a.LEKPR) as MINLEKPR from BWINVI b inner join Artlief a on b.artnr = a.artnr where a.lekpr > 0 group by a.artnr"
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    Do While Not rsrs.EOF
                    
                        dMINLEKPR = 0
                        sArt = "0"
                        If Not IsNull(rsrs!artnr) Then
                            sArt = rsrs!artnr
                        End If
                        
                        If Not IsNull(rsrs!MINLEKPR) Then
                            dMINLEKPR = rsrs!MINLEKPR
                        End If
                        
                        sSQL = "Update BWINVI "
                        sSQL = sSQL & " set LEKPR = '" & dMINLEKPR & "'"
                        sSQL = sSQL & " where artnr = " & sArt
                        gdBase.Execute sSQL, dbFailOnError
                        
                        rsrs.MoveNext
                    Loop
    
                End If
                rsrs.Close: Set rsrs = Nothing

                sSQL = "Update BWINVI "
                sSQL = sSQL & " set LEKWERT =  LEKPR * Bestand "
                gdBase.Execute sSQL, dbFailOnError
            
                reportbildschirm "INVENe", "aWKL71c"
            ElseIf Option1(1).value = True Then
                reportbildschirm "INVENe", "aWKL71b"
            End If
        Case 5 ' leere srechnertab & "ATI aus mde
            iRet = MsgBox("Möchten Sie wirklich löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbYes Then
                delATI
                
                Label7(10).Caption = "0"
                Label7(11).Caption = "0"

                List3.Clear
                Text3(0).SetFocus
            End If
        Case 11
        
            If Check1.value = vbChecked Then 'in Excel kumulieren für Poppitz
                If Text3(1).Text = "" Then Text3(1).Text = "1"
                If Excelupdate(Text3(0).Text, CInt(Text3(1).Text)) = True Then
                    anzeigeNew "normal", "Dieser Artikel wurde erkannt.", Label6
                Else
                    anzeigeNew "rot", "Dieser Artikel wurde nicht erkannt.", Label6
                End If
                
                
            Else
                Dim cValid As String
                Dim cFeld As String
                Dim cZeichen As String
                Dim lcount As Long
                Dim bTextSuche As Boolean
                
                Screen.MousePointer = 11
                
                cValid = "1234567890"
                cFeld = Text3(0).Text
                
                bTextSuche = False
                
                For lcount = 1 To Len(cFeld)
                    cZeichen = Mid(cFeld, lcount, 1)
                    If InStr(cValid, cZeichen) = 0 Then
                        bTextSuche = True
                        Exit For
                    End If
                Next lcount
                
                If bTextSuche Then
                    gcSuch = Text3(0).Text
                    gsARTNR = ""
                    frmWKL70.Show 1
                    Me.Refresh
                    If gsARTNR <> "" Then
                        Text3(0).Text = gsARTNR
                        gsARTNR = ""
                    End If
                End If
                Screen.MousePointer = 0
            
                If Text3(1).Text = "" Then Text3(1).Text = "1"
                
                If artikelgefunden(Text3(0).Text) Then
                    speicherIN Text3(0).Text, CLng(Text3(1).Text)
                    fuellelistlf List3
                    Label7(10).Caption = ermVart(srechnertab & "ATI")
                    Label7(11).Caption = ermGBart(srechnertab & "ATI")
                    If Check9.value = vbChecked Then
                        Text3(1).Text = "1"
                    End If
                    Text3(0).Text = ""
                    Text3(0).SetFocus
                Else
                    If artikelgefundenA1(Text3(0).Text) Then
                        speicherINA1 Text3(0).Text, CLng(Text3(1).Text)
                        fuellelistScan List3
                        Label7(10).Caption = ermVart(srechnertab & "ATI")
                        Label7(11).Caption = ermGBart(srechnertab & "ATI")
                        If Check9.value = vbChecked Then
                            Text3(1).Text = "1"
                        End If
                        Text3(0).Text = ""
                        Text3(0).SetFocus
                    Else
                        Text3(0).SetFocus
                        anzeigeNew "rot", "Dieser Artikel wurde nicht erkannt.", Label6
                    End If
                End If
            End If
        Case 23
            MDElesen
            If mdeErr Then
                reportbildschirm "", "aWKL46e" 'Error artikel mde
            End If
            
        Case 22 'Druckvorschau aus mde
        
            If Not NewTableSuchenDBKombi(srechnertab & "ATO_MDE", gdBase) Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        
        
            loeschNEW "BWINV", gdBase
            CreateTableT2 "BWINV", gdBase
            
            sSQL = "Insert into BWINV select "
            sSQL = sSQL & "  ARTNR "
            sSQL = sSQL & ", BEZEICH "
            sSQL = sSQL & ", AGN "
            sSQL = sSQL & ", LIBESNR "
            sSQL = sSQL & ", BESTAND "
            sSQL = sSQL & ", KVKPR1 "
            sSQL = sSQL & ", VKPR "
            sSQL = sSQL & ", LEKPR "
            sSQL = sSQL & ", EKPR "
            sSQL = sSQL & ", LINR "
            sSQL = sSQL & ", LPZ "
            sSQL = sSQL & ", EAN "
            sSQL = sSQL & ", EAN2 "
            sSQL = sSQL & ", EAN3 "
            sSQL = sSQL & " from " & srechnertab & "ATO_MDE "
            gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update BWINV inner join Artikel "
            sSQL = sSQL & "  on BWINV.ARTNR = Artikel.artnr "
            sSQL = sSQL & " set BWINV.RKZ = Artikel.RKZ "
            sSQL = sSQL & " , BWINV.MWST = Artikel.MWST "
            gdBase.Execute sSQL, dbFailOnError
        
            If Modul6.FindFile(gcDBPfad, "aWKL71as.rpt") Then
                reportbildschirm "INVENe", "aWKL71as"
            Else
                reportbildschirm "INVENe", "aWKL71a"
            End If
        
'            reportbildschirm "INVENe", "aWKL71a"
            
        Case 21 ' leere srechnertab & "ATO aus mde
            iRet = MsgBox("Möchten Sie wirklich löschen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbYes Then
                delATO
                MDEbackanz
            Else
                
            End If
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command2_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Function Excelupdate(sEAN As String, iMenge As Integer) As Boolean
On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cPfad As String
    Dim cDatname As String
    Dim dbExcel As Database
    Dim lAnz As Long
    Dim lDisAnz As Long
    Dim rsrs As Recordset
    Dim gsExcel50 As String
    
    gsExcel50 = "Excel 5.0;"
    
    Screen.MousePointer = 11
    
    Excelupdate = False
    
    lAnz = 0
    lDisAnz = 0

    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "BOX\"
    
    Set dbExcel = OpenDatabase(cPfad & "Masterdatei_Top200_Sep2010_Dtl.xls", 0, 0, gsExcel50)
    
    Set rsrs = dbExcel.OpenRecordset("Formular$")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!EAN) Then
                If Trim(sEAN) = Trim(rsrs!EAN) Then
                    Excelupdate = True
                    rsrs.Edit
                    
                    If Not IsNull(rsrs!Regal) Then
                        rsrs!Regal = rsrs!Regal + iMenge
                        
                    Else
                        rsrs!Regal = iMenge
                    End If
                    
                    rsrs.Update
                
                End If
            End If
            
            If Not IsNull(rsrs!Regal) Then
                If Val(rsrs!Regal) > 0 Then
                    lAnz = lAnz + Val(rsrs!Regal)
                    lDisAnz = lDisAnz + 1
                End If
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0

    Label7(11).Caption = CStr(lAnz)
    Label7(10).Caption = CStr(lDisAnz)
    
    dbExcel.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Excelupdate"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub speicherINA1(sArt As String, lMenge As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    
    sArt = Trim(sArt)
    
    If sArt <> "" Then
        cSQL = "Insert into " & srechnertab & "ATI  select artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", (EKPR * Bestand) as SEKWERT "
        cSQL = cSQL & ", 0 as  LEKPR "
        cSQL = cSQL & ", 0 as LEKWERT "
        cSQL = cSQL & ", KVKPR1 "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", " & lMenge & " as Bestand "
        cSQL = cSQL & ", LIBESNR "
        cSQL = cSQL & " from artikel where artnr = " & sArt
        gdBase.Execute cSQL, dbFailOnError
        
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherINA1"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub fuellelistlf(lst As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    Dim lMax    As Long
    
    cSQL = "Select max(lfnr) as maxi from " & srechnertab & "ATI "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            lMax = rsrs!maxi
        Else
            lMax = 0
        End If
    Else
        lMax = 0
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    cSQL = "Select * from " & srechnertab & "ATI where lfnr =" & lMax
    Set rsrs = gdBase.OpenRecordset(cSQL)
    

    If Not rsrs.EOF Then
        

        If Not IsNull(rsrs!artnr) Then
            cFeld = rsrs!artnr
        End If

        cLBSatz = cFeld & Space(7 - Len(cFeld))
        
        If Not IsNull(rsrs!BEZEICH) Then
            cFeld = rsrs!BEZEICH
        Else
            cFeld = ""
        End If
        
        cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
        
        If Not IsNull(rsrs!BESTAND) Then
            cFeld = rsrs!BESTAND
        Else
            cFeld = ""
        End If
        
        cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))
        
        
        
        
        lst.AddItem cLBSatz, 0
            

    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelistlf"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub speicherIN(sEAN As String, lMenge As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    sEAN = Trim(sEAN)
    
    If sEAN <> "" Then
    
        cSQL = "Insert into " & srechnertab & "ATI select artnr "
        cSQL = cSQL & ", BEZEICH "
        cSQL = cSQL & ", LINR "
        cSQL = cSQL & ", LPZ "
        cSQL = cSQL & ", AGN "
        cSQL = cSQL & ", EAN "
        
        cSQL = cSQL & ", EKPR "
        cSQL = cSQL & ", (EKPR * Bestand) as SEKWERT "
        cSQL = cSQL & ", 0 as  LEKPR "
        cSQL = cSQL & ", 0 as LEKWERT "
        
        cSQL = cSQL & ", KVKPR1 "
        cSQL = cSQL & ", VKPR "
        cSQL = cSQL & ", " & lMenge & " as Bestand "
        cSQL = cSQL & ", LIBESNR "
        
        If Len(sEAN) = 11 Then
            sEAN = "0" & sEAN
    
            cSQL = cSQL & " from artikel where ean = '" & sEAN & "'"
            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        ElseIf Len(sEAN) = 8 Then
            If Left(sEAN, 1) = "2" Then
                sEAN = Mid$(sEAN, 2, 6)
                cSQL = cSQL & " from artikel where artnr = " & sEAN
            Else
                cSQL = cSQL & " from artikel where ean = '" & sEAN & "'"
                cSQL = cSQL & " or ean2 = '" & sEAN & "'"
                cSQL = cSQL & " or ean3 = '" & sEAN & "'"
            End If
        ElseIf Len(sEAN) = 6 Then
            
            cSQL = cSQL & " from artikel where artnr = " & sEAN
        Else
            cSQL = cSQL & " from artikel where ean = '" & sEAN & "'"
            cSQL = cSQL & " or ean2 = '" & sEAN & "'"
            cSQL = cSQL & " or ean3 = '" & sEAN & "'"
        End If
        gdBase.Execute cSQL, dbFailOnError
        

    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherIN"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function artikelgefunden(sEAN As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset
    Dim cPreis      As String
    Dim lLinr       As Long
    Dim dPreis      As Double
    
    artikelgefunden = False
    sEAN = Trim(sEAN)
    
    If sEAN <> "" Then
    
        If Len(sEAN) >= 13 And Left(sEAN, 3) = "419" Then
            dPreis = Val(Mid(sEAN, 9, 4))
            dPreis = dPreis / 100
            lLinr = glZeitungsLinr 'ermLinrInZeitE
            
            If lLinr > 0 Then
                sEAN = ermartnrausLIBESNR(CStr(Val(Mid(sEAN, 4, 5))), lLinr)
                
                If sEAN = "" Then
                    Exit Function
                Else
                    Text3(0).Text = sEAN
                    
                    cSQL = "Update artikel set ekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " ,Lekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " where artnr = " & sEAN
                    gdBase.Execute cSQL, dbFailOnError
                End If
            End If
        End If
        
        If Len(sEAN) >= 13 And Left(sEAN, 3) = "414" Then
            
            dPreis = Val(Mid(sEAN, 9, 4))
            dPreis = dPreis / 100
            lLinr = glZeitungsLinr ' ermLinrInZeitE
            
            If lLinr > 0 Then
                sEAN = ermartnrausLIBESNR(CStr(Val(Mid(sEAN, 4, 5))), lLinr)
                
                If sEAN = "" Then
                    Exit Function
                Else
                    Text3(0).Text = sEAN
                    
                    cSQL = "Update artikel set ekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " ,Lekpr =  '" & dPreis & "'"
                    cSQL = cSQL & " where artnr = " & sEAN
                    gdBase.Execute cSQL, dbFailOnError
                End If
            End If
        End If
    
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
            artikelgefunden = True
            anzeige "normal", "letzter Artikel: " & rsArt!BEZEICH & " Menge: " & Text3(1).Text, Label6
        
        End If
        rsArt.Close: Set rsArt = Nothing
    
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikelgefunden"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function artikelgefundenA1(sArt As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset
    
    artikelgefundenA1 = False
    sArt = Trim(sArt)
    
    If sArt <> "" Then
        If Len(sArt) = 8 Then
            If Left(sArt, 1) = "2" Then
                sArt = Mid(sArt, 2, 6)
            ElseIf Left(sArt, 1) = "0" Then
                sArt = Mid(sArt, 2, 6)
            Else
                sArt = ""
            End If
        Else
'            sart = ""
        End If
    
        If Len(sArt) < 7 And IsNumeric(sArt) Then
            cSQL = "select * from artikel where artnr = " & sArt
            
            Set rsArt = gdBase.OpenRecordset(cSQL)
            If Not rsArt.EOF Then
                Text3(0).Text = sArt
                artikelgefundenA1 = True
                
                anzeigeNew "normal", "letzter Artikel: " & rsArt!BEZEICH & " Menge: " & Text3(1).Text, Label6
            End If
            rsArt.Close: Set rsArt = Nothing
        End If
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "artikelgefundenA1"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub MDEbackanz()
    On Error GoTo LOKAL_ERROR
    
    Label7(22).Caption = "0"
    Label7(21).Caption = "0"

    List4.Clear
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MDEbackanz"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub delATO()
    On Error GoTo LOKAL_ERROR
    
    loeschNEW srechnertab & "ATO_MDE", gdBase
    CreateTableT2 srechnertab & "ATO_MDE", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "delATO"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub delATI()
    On Error GoTo LOKAL_ERROR
    
    
    loeschNEW srechnertab & "ATI", gdBase
    CreateTableT2 srechnertab & "ATI", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "delATI"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MDElesen()
    On Error GoTo LOKAL_ERROR
    
    If MDEeinlesenOhneLinr(Label6, txtStatus, picprogress, frmWKL71) = False Then
        anzeigeNew "rot", "Es konnten keine Daten aus dem MDE - Gerät ausgelesen werden.", Label6
    Else
        anzeigeNew "normal", "", Label6
        MdeVerarbeitung
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MDElesen"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub MdeVerarbeitung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsMDE       As Recordset
    Dim rsFilB      As Recordset
    Dim rsFilBu     As Recordset
    Dim rsArt       As Recordset
    Dim seekEAN     As String
    Dim lMenge      As Long
    Dim lscanfolge  As Long
    
    Screen.MousePointer = 11

    Set rsFilB = gdBase.OpenRecordset(srechnertab & "ATO_MDE")
    Set rsFilBu = gdBase.OpenRecordset("ARTERRIN")
    
    mdeErr = False
    lscanfolge = 0
    
    anzeigeNew "normal", "Die Daten aus dem MDE - Gerät werden verarbeitet...", Label6
    
    Set rsMDE = gdBase.OpenRecordset("mdeinh")
    If Not rsMDE.EOF Then
        rsMDE.MoveFirst
        Do While Not rsMDE.EOF
        
            lscanfolge = lscanfolge + 1
            If Not IsNull(rsMDE!eancode) Then
            
                seekEAN = Trim(rsMDE!eancode)
                seekEAN = checkean(seekEAN)
                
                
                If Len(seekEAN) = 11 Then
                    seekEAN = "0" & seekEAN
            
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                ElseIf Len(seekEAN) = 8 Then
                    If Left(seekEAN, 1) = "2" Then
'                        If gsMDEGERAET = "BELAMDE" Then
'                            sSQL = "select artikel.* from artikel "
'                            sSQL = sSQL & " inner join artlief on artikel.artnr = artlief.artnr where artlief.libesnr = '" & seekEAN & "'"
''                            sSQL = sSQL & " and artlief.linr = " & sLinr
'                        Else
                            seekEAN = Mid$(seekEAN, 2, 6)
                            sSQL = "select * from artikel where artnr = " & seekEAN
'                        End If
                    Else
                        sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                        sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                    End If
                
                Else
                    sSQL = "select * from artikel where ean = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean2 = '" & seekEAN & "'"
                    sSQL = sSQL & " or ean3 = '" & seekEAN & "'"
                End If

                Set rsArt = gdBase.OpenRecordset(sSQL)
                
                If Not rsArt.EOF Then 'hier die bekannten
                    rsFilB.AddNew
                
                    rsFilB!artnr = rsArt!artnr
                    rsFilB!BEZEICH = rsArt!BEZEICH
                    rsFilB!AGN = rsArt!AGN
                    rsFilB!linr = rsArt!linr
                    rsFilB!LIBESNR = rsArt!LIBESNR
                    rsFilB!LPZ = rsArt!LPZ
                    rsFilB!KVKPR1 = rsArt!KVKPR1
                    rsFilB!ekpr = rsArt!ekpr
                    rsFilB!lekpr = rsArt!lekpr
                    rsFilB!BESTAND = rsMDE!Menge
                    rsFilB!EAN = rsArt!EAN
                    rsFilB!EAN2 = rsArt!EAN2
                    rsFilB!EAN3 = rsArt!EAN3
                    rsFilB!vkpr = rsArt!vkpr
                   

                    rsFilB.Update
                Else 'hier die unbekannten
                
                    mdeErr = True
                    rsFilBu.AddNew
                    rsFilBu!EAN = seekEAN
                    rsFilBu!Menge = rsMDE!Menge
                    rsFilBu!lfnr = lscanfolge
                    
                    rsFilBu.Update
                    
                End If
                rsArt.Close: Set rsArt = Nothing
            End If
            rsMDE.MoveNext
        Loop
    
    End If
    
    rsMDE.Close: Set rsMDE = Nothing
    rsFilB.Close: Set rsFilB = Nothing
    rsFilBu.Close: Set rsFilBu = Nothing
    
    fuellelist List4
    Label7(22).Caption = ermVart(srechnertab & "ATO_MDE")
    Label7(21).Caption = ermGBart(srechnertab & "ATO_MDE")
    
    anzeigeNew "normal", "Der Einlesevorgang ist beendet.", Label6
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MdeVerarbeitung"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub fuellelist(lst As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    
    
    cSQL = "Select * from " & srechnertab & "ATO_MDE order by lfnr desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    lst.Clear
    lst.Visible = False
    
    If Not rsrs.EOF Then
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            End If
    
            cLBSatz = cFeld & Space(7 - Len(cFeld))
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            
            If Not IsNull(rsrs!BESTAND) Then
                cFeld = rsrs!BESTAND
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))
            
            
        
            lst.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    lst.Visible = True
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelist"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub fuellelistScan(lst As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim cFeld   As String
    Dim cLBSatz As String
    
    
    cSQL = "Select * from " & srechnertab & "ATI order by lfnr desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    lst.Clear
    lst.Visible = False
    
    If Not rsrs.EOF Then
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cFeld = rsrs!artnr
            End If
    
            cLBSatz = cFeld & Space(7 - Len(cFeld))
            
            If Not IsNull(rsrs!BEZEICH) Then
                cFeld = rsrs!BEZEICH
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(36 - Len(cFeld))
            
            If Not IsNull(rsrs!BESTAND) Then
                cFeld = rsrs!BESTAND
            Else
                cFeld = ""
            End If
            
            cLBSatz = cLBSatz & cFeld & Space(6 - Len(cFeld))
            
            lst.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    lst.Visible = True
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellelistScan"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub

Private Sub Command3_Click()
Frame2.Visible = False
End Sub

Private Sub Command5_Click(index As Integer)
On Error GoTo LOKAL_ERROR

anzeige "normal", "", Label6

Select Case index

    Case 0
        loeschNEW srechnertab & "ATO_MDE", gdBase
        CreateTableT2 srechnertab & "ATO_MDE", gdBase
        
        loeschNEW srechnertab & "ATI", gdBase
        CreateTableT2 srechnertab & "ATI", gdBase
        Unload frmWKL71
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Function ermVart(sTab As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset
    Dim lvanz       As Long
    
    ermVart = "0"
    
    cSQL = "select artnr from " & sTab & " group by artnr"
    Set rsArt = gdBase.OpenRecordset(cSQL)
    If Not rsArt.EOF Then
        rsArt.MoveLast
        lvanz = rsArt.RecordCount
        ermVart = CStr(lvanz)
    End If
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermVart"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function ermGBart(sTab As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsArt       As Recordset

    ermGBart = "0"
    
    cSQL = "select sum(bestand) as lvanz from " & sTab & " "
    Set rsArt = gdBase.OpenRecordset(cSQL)
    If Not rsArt.EOF Then
        If Not IsNull(rsArt!lvanz) Then
            ermGBart = rsArt!lvanz
        End If
    End If
    rsArt.Close: Set rsArt = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermGBart"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub vorbereitungMDE()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    loeschNEW "ARTERRIN", gdBase
    CreateTable "ARTERRIN", gdBase
    
    If Not NewTableSuchenDBKombi(srechnertab & "ATO_MDE", gdBase) Then
        CreateTableT2 srechnertab & "ATO_MDE", gdBase
    End If
    
    If Not NewTableSuchenDBKombi(srechnertab & "ATI", gdBase) Then
        CreateTableT2 srechnertab & "ATI", gdBase
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "vorbereitungMDE"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command7_Click()
On Error GoTo LOKAL_ERROR
    
    Dim iText  As Integer
    iText = CInt(Text3(1).Text)
    If iText = 999 Then
    
    Else
        iText = iText + 1
        Text3(1).Text = CStr(iText)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer1.Enabled = True

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer1.Enabled = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command7_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer2.Enabled = True

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_MouseDown"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR
    Timer2.Enabled = False

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_MouseUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command8_Click()
On Error GoTo LOKAL_ERROR
    
    Dim iText  As Integer
    iText = CInt(Text3(1).Text)
    If iText = -999 Then
    
    Else
        iText = iText - 1
        Text3(1).Text = CStr(iText)
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command8_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

 TextConvPfad.Text = gsConverterPfad

PositionierenWKL71
Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
Modul6.Farbform Me, lblUeberschrift
vorbereitungMDE
anzeige "normal", "", Label6

Timer1.Enabled = False
Timer2.Enabled = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    Label1(20).ForeColor = glS1
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Frame6_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Label1_Click(index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case index
        Case Is = 20 'Inventur mit Scanner
            URLGoTo Me.hwnd, "http://www.kisslive.de/winkiss/hilfe-bei-problemen/44-software-probleme-winkiss/221-aktualisieren-der-artikel-stammdaten-von-mde-geraeten.html"
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_Click"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo LOKAL_ERROR

    If index = 20 Then
        Label1(20).ForeColor = glLink
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label1_MouseMove"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Timer1_Timer()
    On Error GoTo LOKAL_ERROR
    
    Command7_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer1_Timer"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Timer2_Timer()
    On Error GoTo LOKAL_ERROR
    
    Command8_Click
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Timer2_Timer"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub PositionierenWKL71()
On Error GoTo LOKAL_ERROR

    With Frame10
        .Height = 6615
        .Width = 11775
        .Top = 960
        .Left = 0
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame9
        .Height = 6615
        .Width = 11775
        .Top = 960
        .Left = 0
        .BorderStyle = 0
        .Visible = False
    End With
    
    With Frame6
        .Height = 6735
        .Width = 11535
        .Top = 960
        .Left = 120
        .BorderStyle = 0
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenWKL71"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR

    loeschNEW "BWINVI", gdBase
    loeschNEW "BWINV", gdBase
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text3_GotFocus(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Text3(index).BackColor = glSelBack1
    Text3(index).SelStart = 0
    Text3(index).SelLength = Len(Text3(index).Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    If index = 0 And (KeyCode = 187 Or KeyCode = 106) Then
        If Len(Text3(0).Text) > 1 Then
            Text3(1).Text = Left(Text3(0).Text, Len(Text3(0).Text) - 1)
            Text3(0).Text = ""
            Text3(0).SetFocus
        End If
    End If
    
    If KeyCode = vbKeyReturn Then
        If index = 0 Then
            Command2_Click 11
        ElseIf index = 1 Then
            
            Text3(1).Text = ""
            Text3(0).Text = ""
            Command2_Click 11
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        Command2_Click 3
    End If
    
    If KeyCode = vbKeyRight Then
        Text3(0).SetFocus
    End If
    
    If KeyCode = vbKeyLeft Then
        Text3(1).SetFocus
    End If
    
    If index = 1 Then
        If KeyCode = vbKeyUp Then
            Text3(1).Text = CInt(Text3(1).Text) + 1
        End If
        
        If KeyCode = vbKeyDown Then
            Text3(1).Text = CInt(Text3(1).Text) - 1
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text3_LostFocus(index As Integer)
On Error GoTo LOKAL_ERROR

    Text3(index).BackColor = vbWhite
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil Artikelliste aus MDE /Scanner ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
