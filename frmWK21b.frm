VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWK21b 
   Caption         =   "Tagesabschluss/Kasseninhalt"
   ClientHeight    =   8595
   ClientLeft      =   1485
   ClientTop       =   1830
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "frmWK21b.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'Kein
      Height          =   1335
      Left            =   0
      TabIndex        =   71
      Top             =   7680
      Width           =   12015
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   4
         Left            =   9720
         TabIndex        =   85
         Top             =   120
         Visible         =   0   'False
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
         Caption         =   "Weiter"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   75
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
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
         Caption         =   "Suchen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   2
         Left            =   3720
         TabIndex        =   74
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
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
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
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
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   0
         Left            =   7800
         TabIndex        =   72
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
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
         Caption         =   "Schlie?en"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   615
         Index           =   5
         Left            =   5400
         TabIndex        =   93
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
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
         Caption         =   "Terminalschnitt"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'Kein
      Height          =   975
      Left            =   0
      TabIndex        =   69
      Top             =   6720
      Width           =   12135
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "1"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   1
         Left            =   960
         TabIndex        =   95
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "2"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   2
         Left            =   1800
         TabIndex        =   96
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "3"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   3
         Left            =   2640
         TabIndex        =   97
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "4"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   4
         Left            =   3480
         TabIndex        =   98
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "5"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   5
         Left            =   4320
         TabIndex        =   99
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "6"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   6
         Left            =   5160
         TabIndex        =   100
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "7"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   7
         Left            =   6000
         TabIndex        =   101
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "8"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   8
         Left            =   6840
         TabIndex        =   102
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "9"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   9
         Left            =   7680
         TabIndex        =   103
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "0"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   10
         Left            =   8520
         TabIndex        =   104
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   ","
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   11
         Left            =   9360
         TabIndex        =   105
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "C"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   12
         Left            =   10200
         TabIndex        =   106
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "<<"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   855
         Index           =   13
         Left            =   11040
         TabIndex        =   107
         Top             =   0
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   ">>"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Label Label0 
         BackColor       =   &H00FF0000&
         Caption         =   "-1"
         Height          =   375
         Left            =   8760
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   8280
      TabIndex        =   65
      Top             =   0
      Width           =   3735
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   27
         Left            =   1680
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bargeld :"
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
         Index           =   51
         Left            =   120
         TabIndex        =   90
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "123456,89"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   1680
         TabIndex        =   67
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Summe :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'Kein
      Height          =   5775
      Left            =   0
      TabIndex        =   30
      Top             =   840
      Width           =   12015
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   28
         Left            =   9960
         TabIndex        =   91
         Text            =   "Text1"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   25
         Left            =   9960
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   24
         Left            =   9960
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   23
         Left            =   9960
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   22
         Left            =   9960
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   21
         Left            =   9960
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   20
         Left            =   9960
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   19
         Left            =   9960
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   18
         Left            =   9960
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   17
         Left            =   4800
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   16
         Left            =   4800
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   15
         Left            =   4800
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   14
         Left            =   4800
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   13
         Left            =   4800
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   12
         Left            =   4800
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   11
         Left            =   4800
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   10
         Left            =   4800
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   9
         Left            =   1440
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   8
         Left            =   1440
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   1440
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   1440
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Differenz:"
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
         Index           =   52
         Left            =   6960
         TabIndex        =   92
         Top             =   5160
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Dukaten Bestand:"
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
         Index           =   50
         Left            =   6960
         TabIndex        =   89
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   49
         Left            =   240
         TabIndex        =   88
         Top             =   4560
         Width           =   6615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   3555
         X2              =   3555
         Y1              =   240
         Y2              =   4440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   0
         X1              =   6910
         X2              =   6910
         Y1              =   240
         Y2              =   5160
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
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
         Index           =   48
         Left            =   9960
         TabIndex        =   87
         Top             =   4680
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Wechselgeld - wird vorgetragen f?r den n?chsten Tag:"
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
         Index           =   47
         Left            =   6960
         TabIndex        =   86
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "zur Bank - Absch?pfungsbetrag:"
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
         Index           =   46
         Left            =   6960
         TabIndex        =   84
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kreditkarte St?ck:"
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
         Index           =   45
         Left            =   6960
         TabIndex        =   83
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kreditkarte Wert:"
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
         Index           =   44
         Left            =   6960
         TabIndex        =   82
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gutscheine St?ck:"
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
         Index           =   43
         Left            =   6960
         TabIndex        =   81
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Gutscheine Wert:"
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
         Index           =   42
         Left            =   6960
         TabIndex        =   80
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Dukaten St?ck:"
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
         Index           =   41
         Left            =   6960
         TabIndex        =   79
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Dukaten Wert:"
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
         Index           =   40
         Left            =   6960
         TabIndex        =   78
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "verschiedenes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   6960
         TabIndex        =   77
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   35
         Left            =   5760
         TabIndex        =   64
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   34
         Left            =   5760
         TabIndex        =   63
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   33
         Left            =   5760
         TabIndex        =   62
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   32
         Left            =   5760
         TabIndex        =   61
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   31
         Left            =   5760
         TabIndex        =   60
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   30
         Left            =   5760
         TabIndex        =   59
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   29
         Left            =   5760
         TabIndex        =   58
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   28
         Left            =   5760
         TabIndex        =   57
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   27
         Left            =   2400
         TabIndex        =   56
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   26
         Left            =   2400
         TabIndex        =   55
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   25
         Left            =   2400
         TabIndex        =   54
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   24
         Left            =   2400
         TabIndex        =   53
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   23
         Left            =   2400
         TabIndex        =   52
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   22
         Left            =   2400
         TabIndex        =   51
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   21
         Left            =   2400
         TabIndex        =   50
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00C0C000&
         Caption         =   "0"
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
         Index           =   20
         Left            =   2400
         TabIndex        =   49
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "0,01er:"
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
         Index           =   19
         Left            =   3600
         TabIndex        =   48
         Tag             =   "0,01"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "0,02er:"
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
         Index           =   18
         Left            =   3600
         TabIndex        =   47
         Tag             =   "0,02"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "0,05er:"
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
         Index           =   17
         Left            =   3600
         TabIndex        =   46
         Tag             =   "0,05"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "0,10er:"
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
         Index           =   16
         Left            =   3600
         TabIndex        =   45
         Tag             =   "0,10"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "0,20er:"
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
         Index           =   15
         Left            =   3600
         TabIndex        =   44
         Tag             =   "0,20"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "0,50er:"
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
         Index           =   14
         Left            =   3600
         TabIndex        =   43
         Tag             =   "0,50"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "1er:"
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
         Index           =   13
         Left            =   3600
         TabIndex        =   42
         Tag             =   "1"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "2er:"
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
         Index           =   12
         Left            =   3600
         TabIndex        =   41
         Tag             =   "2"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "5er:"
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
         Index           =   11
         Left            =   240
         TabIndex        =   40
         Tag             =   "5"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "10er:"
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
         Index           =   10
         Left            =   240
         TabIndex        =   39
         Tag             =   "10"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "20er:"
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
         Index           =   9
         Left            =   240
         TabIndex        =   38
         Tag             =   "20"
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "50er:"
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
         Index           =   8
         Left            =   240
         TabIndex        =   37
         Tag             =   "50"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "100er:"
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
         Index           =   7
         Left            =   240
         TabIndex        =   36
         Tag             =   "100"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "200er:"
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
         Index           =   6
         Left            =   240
         TabIndex        =   35
         Tag             =   "200"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "500er:"
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
         Index           =   5
         Left            =   240
         TabIndex        =   34
         Tag             =   "500"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "1000er:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   33
         Tag             =   "1000"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Anzahl M?nzen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   32
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Anzahl Scheine:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   8295
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   4080
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   26
         Left            =   7080
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1560
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "W?hrung:"
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
         Index           =   38
         Left            =   5760
         TabIndex        =   68
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Datum:"
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
         Index           =   1
         Left            =   3000
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Filiale:"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmWK21b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub DruckDatenBargeldWK21b()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim cDatum As String
    Dim dDatum As Double
    Dim cFiliale As String
    
    Dim cTag As String
    Dim cMon As String
    Dim cJahr As String
    
    Dim cDrucker As String
    Dim bReturn As Boolean
    Dim lAnz As Long
    Dim lcount As Long
    
    setzedrucker gcBonDrucker

    loeschNEW "DRU_BARG", gdBase
    CreateTable "DRU_BARG", gdBase
    
    cFiliale = Text1(0).Text
    cFiliale = Trim$(cFiliale)
    If cFiliale = "" Then
        cFiliale = "1"
    End If
    
    dDatum = 0
    cDatum = Text1(1).Text
    cDatum = Trim$(cDatum)
    cTag = Day(cDatum)
    cMon = Month(cDatum)
    cJahr = Year(cDatum)
    dDatum = DateSerial(cJahr, cMon, cTag)
    
    cSQL = "Insert into DRU_BARG select * from BARGELD "
    cSQL = cSQL & "where BARGELD.FILIALE = " & cFiliale & " "
    cSQL = cSQL & "and BARGELD.DATUM = " & Trim$(Str$(dDatum)) & " "
    gdBase.Execute cSQL, dbFailOnError
    
    SendeDaten2DruckerNeuWKL21b
    
    setzedrucker gcListenDrucker

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckDatenBargeldWK21b"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
End Sub
Private Sub DruckDatenBargeldWK21b_DINA4(bPrinter_direkt As Boolean)
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cTmp2 As String
    Dim cWaeCode As String
    Dim cDatum As String
    Dim cFiliale As String
    
    cFiliale = Text1(0).Text
    cFiliale = Trim$(cFiliale)
    If cFiliale = "" Then
        cFiliale = "1"
    End If
    
    cDatum = Text1(1).Text
    
    '//new
    If gcWaehrung = "EUR" Then
        cWaeCode = "alle Preise in EURO"
    Else
        cWaeCode = "alle Preise in " & gcWaehrung
    End If

    loeschNEW "DRU_BARG_DIN", gdBase

    cSQL = "Create Table DRU_BARG_DIN "
    cSQL = cSQL & "("
    cSQL = cSQL & " WAE_CODE TEXT(20)"
    cSQL = cSQL & ", DATEN1 TEXT(50)"
    cSQL = cSQL & ", DATEN2 TEXT(50)"
    cSQL = cSQL & ", DATEN3 TEXT(50)"
    cSQL = cSQL & ", DATEN4 TEXT(50)"
    cSQL = cSQL & ", DATEN5 TEXT(50)"
    cSQL = cSQL & ", DATEN6 TEXT(50)"
    cSQL = cSQL & ", DATEN7 TEXT(50)"
    cSQL = cSQL & ", DATEN8 TEXT(50)"
    cSQL = cSQL & ", DATEN9 TEXT(50)"
    cSQL = cSQL & ", DATEN10 TEXT(50)"
    cSQL = cSQL & ", DATEN11 TEXT(50)"
    cSQL = cSQL & ", DATEN12 TEXT(50)"
    cSQL = cSQL & ", DATEN13 TEXT(50)"
    cSQL = cSQL & ", DATEN14 TEXT(50)"
    cSQL = cSQL & ", DATEN15 TEXT(50)"
    cSQL = cSQL & ", DATEN16 TEXT(50)"
    cSQL = cSQL & ", DATEN17 TEXT(50)"
    cSQL = cSQL & ", DATEN18 TEXT(50)"
    cSQL = cSQL & ", DATEN19 TEXT(50)"
    cSQL = cSQL & ", DATEN20 TEXT(50)"
    cSQL = cSQL & ", DATEN21 TEXT(50)"
    cSQL = cSQL & ", DATEN22 TEXT(50)"
    cSQL = cSQL & ", DATEN23 TEXT(50)"
    cSQL = cSQL & ", DATEN24 TEXT(50)"
    cSQL = cSQL & ", DATEN25 TEXT(50)"
    cSQL = cSQL & ", DATEN26 TEXT(50)"
    cSQL = cSQL & ", DATEN27 TEXT(50)"
    cSQL = cSQL & ", DATEN28 TEXT(50)"
    cSQL = cSQL & ", DATEN29 TEXT(50)"
    cSQL = cSQL & ", DATEN30 TEXT(50)"
    cSQL = cSQL & ", DATEN31 TEXT(50)"
    cSQL = cSQL & ", DATEN32 TEXT(50)"
    cSQL = cSQL & ", DATEN33 TEXT(50)"
    cSQL = cSQL & ", DATEN34 TEXT(50)"
    cSQL = cSQL & ", DATEN35 TEXT(50)"
    cSQL = cSQL & ", DATEN36 TEXT(50)"
    cSQL = cSQL & ", DATEN37 TEXT(50)"
    cSQL = cSQL & ", DATEN38 TEXT(50)"
    cSQL = cSQL & ", DATEN39 TEXT(50)"
    cSQL = cSQL & ", DATEN40 TEXT(50)"
    cSQL = cSQL & ", DATEN41 TEXT(50)"
    cSQL = cSQL & ", DATEN42 TEXT(50)"
    cSQL = cSQL & ", DATEN43 TEXT(50)"
    cSQL = cSQL & ", DATEN44 TEXT(50)"
    cSQL = cSQL & ", DATEN45 TEXT(50)"
    cSQL = cSQL & ", DATEN46 TEXT(50)"
    cSQL = cSQL & ", DATEN47 TEXT(50)"
    cSQL = cSQL & ", DATEN48 TEXT(50)"
    cSQL = cSQL & ", DATEN49 TEXT(50)"
    cSQL = cSQL & ", DATEN50 TEXT(50)"
    cSQL = cSQL & ", DATEN51 TEXT(50)"
    cSQL = cSQL & ", DATEN52 TEXT(50)"
    cSQL = cSQL & ", DATEN53 TEXT(50)"
    cSQL = cSQL & ", DATEN54 TEXT(50)"
    cSQL = cSQL & ", DATEN55 TEXT(50)"
    cSQL = cSQL & ", DATEN56 TEXT(50)"
    cSQL = cSQL & ", DATEN57 TEXT(50)"
    cSQL = cSQL & ", DATEN58 TEXT(50)"
    cSQL = cSQL & ", DATEN59 TEXT(50)"
    cSQL = cSQL & ", DATEN60 TEXT(50)"
    
    cSQL = cSQL & ")"

    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Select * from DRU_BARG_DIN"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If rsrs.EOF Then
        rsrs.AddNew
        
        rsrs!WAE_CODE = cWaeCode

        '**************************************
        ' linke Seite der Kopfdaten
        '**************************************

        '500
        cTmp2 = Text1(3).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN1 = cTmp2
        
        'Wert
        rsrs!DATEN18 = Trim$(Label1(21).Caption)

        '200
        cTmp2 = Text1(4).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN2 = cTmp2
        
        'Wert
        rsrs!DATEN19 = Trim$(Label1(22).Caption)

        '100
        cTmp2 = Text1(5).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN3 = cTmp2
        
        'Wert
        rsrs!DATEN20 = Trim$(Label1(23).Caption)

        '50
        cTmp2 = Text1(6).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN4 = cTmp2
        
        'Wert
        rsrs!DATEN21 = Trim$(Label1(24).Caption)

        '20
        cTmp2 = Text1(7).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN5 = cTmp2
        
        'Wert
        rsrs!DATEN22 = Trim$(Label1(25).Caption)

        '10
        cTmp2 = Text1(8).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN6 = cTmp2
        
        'Wert
        rsrs!DATEN23 = Trim$(Label1(26).Caption)

        '5.00
        cTmp2 = Text1(9).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN7 = cTmp2
        
        'Wert
        rsrs!DATEN24 = Trim$(Label1(27).Caption)
        
        '2.00
        cTmp2 = Text1(10).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN8 = cTmp2
        
        'Wert
        rsrs!DATEN25 = Trim$(Label1(28).Caption)

        '1.00
        cTmp2 = Text1(11).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN9 = cTmp2
        
        'Wert
        rsrs!DATEN26 = Trim$(Label1(29).Caption)

        '0.50
        cTmp2 = Text1(12).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN10 = cTmp2
        
        'Wert
        rsrs!DATEN27 = Trim$(Label1(30).Caption)

        '0.20
        cTmp2 = Text1(13).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN11 = cTmp2
        
        'Wert
        rsrs!DATEN28 = Trim$(Label1(31).Caption)

        '0.10
        cTmp2 = Text1(14).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN12 = cTmp2
        
        'Wert
        rsrs!DATEN29 = Trim$(Label1(32).Caption)
        
        '0.05
        cTmp2 = Text1(15).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN13 = cTmp2
        
        'Wert
        rsrs!DATEN30 = Trim$(Label1(33).Caption)

        '0.02
        cTmp2 = Text1(16).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN14 = cTmp2
        
        'Wert
        rsrs!DATEN31 = Trim$(Label1(34).Caption)


        '0.01
        cTmp2 = Text1(17).Text
        cTmp2 = Trim$(cTmp2)
        rsrs!DATEN15 = cTmp2
        
        'Wert
        rsrs!DATEN32 = Trim$(Label1(35).Caption)

        

        'Kassennummer
        rsrs!DATEN46 = gcKasNum
        
        'Filiale
        rsrs!DATEN47 = cFiliale
        
        'Datum
        rsrs!DATEN48 = cDatum & " / " & Format$(Now, "HH:MM:SS")
        
        'Gesamtsumme
        rsrs!DATEN49 = Trim$(Label1(37).Caption)
        
        'Absch?pfung
        rsrs!DATEN50 = Format(Trim$(Text1(24).Text), "######,##0.00")
        
        'Wechselgeld
        rsrs!DATEN51 = Trim$(Label1(48).Caption)
        
        'Dukaten wert
        rsrs!DATEN52 = Format(Trim$(Text1(18).Text), "######,##0.00")
        
        'Dukaten st?ck
        rsrs!DATEN53 = Trim$(Text1(19).Text)
        
        'Dukaten bestand
        rsrs!DATEN54 = Trim$(Text1(25).Text)
        
        'Gutscheine wert
        rsrs!DATEN55 = Format(Trim$(Text1(20).Text), "######,##0.00")
        
        'Gutscheine st?ck
        rsrs!DATEN56 = Trim$(Text1(21).Text)
        
        'Kreditkarte wert
        rsrs!DATEN57 = Format(Trim$(Text1(22).Text), "######,##0.00")
        
        'Kreditkarte st?ck
        rsrs!DATEN58 = Trim$(Text1(23).Text)
        
          
        'Kassensoll
        rsrs!DATEN59 = Format(ermaktKassensoll, "######,##0.00 ")
        
        'Wechselgeld
        rsrs!DATEN60 = Format(ermaktWechselgeld, "######,##0.00 ")

        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing

    
    If bPrinter_direkt Then
        reportbildschirmToPrinter "aWKL21i"
    Else
        reportbildschirm "", "aWKL21i"
    End If
    
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DruckDatenBargeldWK21b_DINA4"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub LadeBargeldWK21b()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFiliale As String
    Dim cDatum As String
    Dim dDatum As Double
    Dim cWaehrung As String
    Dim iCount As Integer
    Dim dNennwert As Double
    Dim lAnzahl As String
    
    cFiliale = Text1(0).Text
    cFiliale = Trim$(cFiliale)
    If cFiliale = "" Then
        cFiliale = "1"
    End If
    
    cDatum = Text1(1).Text
    cDatum = Trim$(cDatum)
    If Not IsDate(cDatum) Then
        cDatum = Format$(Now, "DD.MM.YYYY")
        Text1(1).Text = cDatum
    End If
    dDatum = DateValue(cDatum)
    
    
    cWaehrung = Text1(26).Text
    cWaehrung = Trim$(cWaehrung)
    If cWaehrung = "" Then
        cWaehrung = gcWaehrung
    End If
    
    cSQL = "Select * from BARGELD "
    cSQL = cSQL & "where FILIALE = " & cFiliale & " "
    cSQL = cSQL & "and DATUM = " & Trim$(Str$(dDatum)) & " "
    cSQL = cSQL & "and WAEHRUNG = '" & cWaehrung & "' "
    cSQL = cSQL & "order by NENNWERT desc "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!NENNWERT) Then
                dNennwert = rsrs!NENNWERT
            Else
                dNennwert = 0
            End If
            
            If Not IsNull(rsrs!ANZAHL) Then
                lAnzahl = rsrs!ANZAHL
            Else
                lAnzahl = "0"
            End If
            
            Select Case dNennwert
                Case Is = 1000
                    '//new
                    If gcWaehrung = "EUR" Then
                    Else
                        Text1(2).Text = lAnzahl
                    End If
                Case Is = 500
                    Text1(3).Text = lAnzahl
                Case Is = 200
                    Text1(4).Text = lAnzahl
                Case Is = 100
                    Text1(5).Text = lAnzahl
                Case Is = 50
                    Text1(6).Text = lAnzahl
                Case Is = 20
                    Text1(7).Text = lAnzahl
                Case Is = 10
                    Text1(8).Text = lAnzahl
                Case Is = 5
                    If Not IsNull(rsrs!art) Then
                        If rsrs!art = "S" Then
                            Text1(9).Text = lAnzahl
                        Else '//rsRs!ART = "M", M?nze
                            '//new
                            If gcWaehrung = "EUR" Then
                            Else
                                Text1(10).Text = lAnzahl
                            End If
                        End If
                    End If
                Case Is = 2
                    '//new
                    If gcWaehrung = "EUR" Then
                        Text1(10).Text = lAnzahl
                    Else
                        Text1(11).Text = lAnzahl
                    End If
                Case Is = 1
                    '//new
                    If gcWaehrung = "EUR" Then
                        Text1(11).Text = lAnzahl
                    Else
                        Text1(12).Text = lAnzahl
                    End If
                Case Is = 0.5
                    '//new
                    If gcWaehrung = "EUR" Then
                        Text1(12).Text = lAnzahl
                    Else
                        Text1(13).Text = lAnzahl
                    End If
                Case Is = 0.2
                    '//new
                    If gcWaehrung = "EUR" Then
                        Text1(13).Text = lAnzahl
                    Else
                    End If
                Case Is = 0.1
                    Text1(14).Text = lAnzahl
                Case Is = 0.05
                    Text1(15).Text = lAnzahl
                Case Is = 0.02
                    Text1(16).Text = lAnzahl
                Case Is = 0.01
                    Text1(17).Text = lAnzahl
                Case Is = 3001
                    Text1(18).Text = lAnzahl
                Case Is = 3002
                    Text1(19).Text = lAnzahl
                Case Is = 3003
                    Text1(20).Text = lAnzahl
                Case Is = 3004
                    Text1(21).Text = lAnzahl
                Case Is = 3005
                    Text1(22).Text = lAnzahl
                Case Is = 3006
                    Text1(23).Text = lAnzahl
                Case Is = 3007
                    Text1(24).Text = lAnzahl
                Case Is = 3008
                    Text1(25).Text = lAnzahl
            End Select
            rsrs.MoveNext
        Loop
    Else
        For iCount = 2 To 25
            Text1(iCount).Text = ""
        Next iCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
    For iCount = 2 To 25
        Text1_LostFocus iCount
    Next iCount
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LadeBargeldWK21b"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
    
    
End Sub
Private Function SchreibeDatenWK21b() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFiliale As String
    Dim cDatum As String
    Dim dDatum As Double
    Dim cWaehrung As String
    Dim cART As String
    Dim cNennWert As String
    Dim dNennwert As Double
    Dim cAnzahl As String
    
    Dim iCount As Integer
    
    SchreibeDatenWK21b = False
    
    cFiliale = Text1(0).Text
    cFiliale = Trim$(cFiliale)
    If cFiliale = "" Then
        cFiliale = "1"
    End If
    
    If CheckDatum(Text1(1).Text) = False Then
        MsgBox "Bitte geben Sie ein Datum ein!", vbInformation, "Winkiss Hinweis:"
        Text1(1).Text = ""
        Text1(1).SetFocus
        Exit Function
    End If
    
    cDatum = Text1(1).Text
    dDatum = DateValue(cDatum)
    dDatum = Fix(dDatum)
    
    
    cWaehrung = Text1(26).Text
    cWaehrung = Trim$(cWaehrung)
    If cWaehrung = "" Then
        cWaehrung = gcWaehrung
    End If
    
    For iCount = 2 To 25
        cAnzahl = Text1(iCount).Text
        cAnzahl = Trim$(cAnzahl)
        If cAnzahl = "" Then
            cAnzahl = "0"
        End If
        
        Select Case iCount
            Case Is = 2
                cART = "S"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = 0
                Else
                    cNennWert = "1000"
                End If
            Case Is = 3
                cART = "S"
                cNennWert = "500"
            Case Is = 4
                cART = "S"
                cNennWert = "200"
            Case Is = 5
                cART = "S"
                cNennWert = "100"
            Case Is = 6
                cART = "S"
                cNennWert = "50"
            Case Is = 7
                cART = "S"
                cNennWert = "20"
            Case Is = 8
                cART = "S"
                cNennWert = "10"
            Case Is = 9
                cART = "S"
                cNennWert = "5"
            Case Is = 10
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "2"
                Else
                    cNennWert = "5"
                End If
            Case Is = 11
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "1"
                Else
                    cNennWert = "2"
                End If
            Case Is = 12
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "0,50"
                Else
                    cNennWert = "1"
                End If
            Case Is = 13
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "0,20"
                Else
                    cNennWert = "0,50"
                End If
            Case Is = 14
                cART = "M"
                cNennWert = "0,10"
            Case Is = 15
                cART = "M"
                cNennWert = "0,05"
            Case Is = 16
                cART = "M"
                cNennWert = "0,02"
            Case Is = 17
                cART = "M"
                cNennWert = "0,01"
            Case Is = 18
                cART = "D"
                cNennWert = "3001"
            Case Is = 19
                cART = "D"
                cNennWert = "3002"
            Case Is = 20
                cART = "G"
                cNennWert = "3003"
            Case Is = 21
                cART = "G"
                cNennWert = "3004"
            Case Is = 22
                cART = "K"
                cNennWert = "3005"
            Case Is = 23
                cART = "K"
                cNennWert = "3006"
            Case Is = 24
                cART = "A"
                cNennWert = "3007"
            Case Is = 25
                cART = "D"
                cNennWert = "3008"
            
        End Select
        
        
        cNennWert = fnMoveComma2Point$(cNennWert)
        dNennwert = Val(cNennWert)
        
        cSQL = "Select * from BARGELD "
        cSQL = cSQL & " where FILIALE = " & cFiliale & " "
        cSQL = cSQL & " and DATUM = " & Trim$(Str$(dDatum)) & " "
        cSQL = cSQL & " and WAEHRUNG = '" & cWaehrung & "' "
        cSQL = cSQL & " and ART = '" & cART & "' "
        cSQL = cSQL & " and NENNWERT = " & cNennWert & " "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            rsrs.Edit
        Else
            rsrs.AddNew
        End If
        
        rsrs!FILIALE = cFiliale
        rsrs!Datum = dDatum
        rsrs!Waehrung = cWaehrung
        rsrs!art = cART
        rsrs!NENNWERT = dNennwert
        rsrs!SENDOK = False
        rsrs!ANZAHL = cAnzahl
        rsrs.Update
        rsrs.Close: Set rsrs = Nothing
        
   Next iCount
   
   SchreibeDatenWK21b = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWK21b"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1

End Function
Private Function SchreibeDatenWK21b_neuDINA4() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cFiliale As String
    Dim cDatum As String
    Dim dDatum As Double
    Dim cWaehrung As String
    Dim cART As String
    Dim cNennWert As String
    Dim dNennwert As Double
    Dim cAnzahl As String
    
    Dim iCount As Integer
    
    SchreibeDatenWK21b_neuDINA4 = False
    
    cFiliale = Text1(0).Text
    cFiliale = Trim$(cFiliale)
    If cFiliale = "" Then
        cFiliale = "1"
    End If
    
    If CheckDatum(Text1(1).Text) = False Then
        MsgBox "Bitte geben Sie ein Datum ein!", vbInformation, "Winkiss Hinweis:"
        Text1(1).Text = ""
        Text1(1).SetFocus
        Exit Function
    End If
    
    cDatum = Text1(1).Text
    dDatum = DateValue(cDatum)
    dDatum = Fix(dDatum)
    
    
    cWaehrung = Text1(26).Text
    cWaehrung = Trim$(cWaehrung)
    If cWaehrung = "" Then
        cWaehrung = gcWaehrung
    End If
    
    For iCount = 2 To 25
        cAnzahl = Text1(iCount).Text
        cAnzahl = Trim$(cAnzahl)
        If cAnzahl = "" Then
            cAnzahl = "0"
        End If
        
        Select Case iCount
            Case Is = 2
                cART = "S"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = 0
                Else
                    cNennWert = "1000"
                End If
            Case Is = 3
                cART = "S"
                cNennWert = "500"
            Case Is = 4
                cART = "S"
                cNennWert = "200"
            Case Is = 5
                cART = "S"
                cNennWert = "100"
            Case Is = 6
                cART = "S"
                cNennWert = "50"
            Case Is = 7
                cART = "S"
                cNennWert = "20"
            Case Is = 8
                cART = "S"
                cNennWert = "10"
            Case Is = 9
                cART = "S"
                cNennWert = "5"
            Case Is = 10
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "2"
                Else
                    cNennWert = "5"
                End If
            Case Is = 11
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "1"
                Else
                    cNennWert = "2"
                End If
            Case Is = 12
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "0,50"
                Else
                    cNennWert = "1"
                End If
            Case Is = 13
                cART = "M"
                '//new
                If gcWaehrung = "EUR" Then
                    cNennWert = "0,20"
                Else
                    cNennWert = "0,50"
                End If
            Case Is = 14
                cART = "M"
                cNennWert = "0,10"
            Case Is = 15
                cART = "M"
                cNennWert = "0,05"
            Case Is = 16
                cART = "M"
                cNennWert = "0,02"
            Case Is = 17
                cART = "M"
                cNennWert = "0,01"
            Case Is = 18
                cART = "D"
                cNennWert = "3001"
            Case Is = 19
                cART = "D"
                cNennWert = "3002"
            Case Is = 20
                cART = "G"
                cNennWert = "3003"
            Case Is = 21
                cART = "G"
                cNennWert = "3004"
            Case Is = 22
                cART = "K"
                cNennWert = "3005"
            Case Is = 23
                cART = "K"
                cNennWert = "3006"
            Case Is = 24
                cART = "A"
                cNennWert = "3007"
            Case Is = 25
                cART = "D"
                cNennWert = "3008"
            
        End Select
        
        
        cNennWert = fnMoveComma2Point$(cNennWert)
        dNennwert = Val(cNennWert)
        
        cSQL = "Select * from BARGELD "
        cSQL = cSQL & " where FILIALE = " & cFiliale & " "
        cSQL = cSQL & " and DATUM = " & Trim$(Str$(dDatum)) & " "
        cSQL = cSQL & " and WAEHRUNG = '" & cWaehrung & "' "
        cSQL = cSQL & " and ART = '" & cART & "' "
        cSQL = cSQL & " and NENNWERT = " & cNennWert & " "
        
        Set rsrs = gdBase.OpenRecordset(cSQL)
        
        If Not rsrs.EOF Then
            rsrs.Edit
        Else
            rsrs.AddNew
        End If
        
        rsrs!FILIALE = cFiliale
        rsrs!Datum = dDatum
        rsrs!Waehrung = cWaehrung
        rsrs!art = cART
        rsrs!NENNWERT = dNennwert
        rsrs!SENDOK = False
        rsrs!ANZAHL = cAnzahl
        rsrs.Update
        rsrs.Close: Set rsrs = Nothing
        
   Next iCount
   
   SchreibeDatenWK21b_neuDINA4 = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeDatenWK21b_neuDINA4"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1

End Function

Private Sub SSCommand2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Set WshShell = CreateObject("WScript.Shell")
'    WshShell.SendKeys "+{Tab}", True
    
    Select Case index
        Case Is < 11
            If index = 10 And (Val(Label0.Caption) = 18 Or Val(Label0.Caption) = 27 Or Val(Label0.Caption) = 20 Or Val(Label0.Caption) = 22 Or Val(Label0.Caption) = 24) Then
                Text1(Label0.Caption).Text = Text1(Label0.Caption).Text & SSCommand2(index).Caption
            ElseIf index <> 10 Then
                Text1(Label0.Caption).Text = Text1(Label0.Caption).Text & SSCommand2(index).Caption
           
            End If
        Case Is = 11
            Text1(Label0.Caption).Text = ""
        Case Is = 12
            WshShell.SendKeys "+{Tab}", True
'            SendKeys "+{TAB}"
        Case Is = 13
            WshShell.SendKeys "{Tab}", True
'            SendKeys "{TAB}"
    End Select
    Text1(Label0.Caption).SetFocus
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 5 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SSCommand2_Click"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
        
        Fehlermeldung1
    End If
End Sub
Private Sub Command1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim dABSCHOPF           As Double
    Dim lDuBestand          As Long
    Dim iRet                As Integer
    Dim dDifferenzbargeld   As Double
    Dim ctmp                As String
    Dim lKJADate            As Long
    Dim cKJAZeit            As String
    Dim sDifftext           As String
    
    Select Case index
        Case Is = 0         'Schlie?en
            
            iRet = MsgBox("Wollen Sie wirklich abbrechen?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
            If iRet = vbNo Then
            
            Else
                Unload frmWK21b
            End If
            
        Case Is = 1         'Speichern
            Screen.MousePointer = 11
            If SchreibeDatenWK21b Then
                
            End If
        
        Case Is = 2         'Drucken
            Screen.MousePointer = 11
            If SchreibeDatenWK21b Then
            
                iRet = MsgBox("M?chten Sie die Kassenabrechnung auf dem Bondrucker drucken?", vbQuestion + vbYesNo + vbDefaultButton2, "Winkiss Frage:")
                
                If iRet = vbYes Then
                    DruckDatenBargeldWK21b
                Else
                    DruckDatenBargeldWK21b_DINA4 False
                End If
            End If
            
        Case Is = 3         'Suchen
            Screen.MousePointer = 11
            LadeBargeldWK21b
            Text1(3).SetFocus
            
        Case Is = 4         'Weiter
            
        
            If Text1(24).Text = "" Then
                anzeige "Rot", "Bitte geben Sie einen Absch?pfungsbetrag ein!", Label1(49)
                Text1(24).SetFocus
                Exit Sub
            End If
            
            
            
            If Label1(37).Caption = "" Then Label1(37).Caption = "0"
            
            If IsNumeric(Label1(37).Caption) = False Then Label1(37).Caption = "0"
            
            dDifferenzbargeld = CDbl(Format(Label1(37).Caption, "########0.00")) - CDbl(Format(ermaktKassensoll, "########0.00"))
            
            dDifferenzbargeld = Format(dDifferenzbargeld, "########0.00")
            
            'Protokoll schreiben
            schreibeProtokoll_Bargeld_Handling "Sie haben " & Format(Label1(37).Caption, "########0.00") & " Euro gez?hlt. Kasse:" & gcKasNum
           
            If dDifferenzbargeld <> 0 Then
            
                schreibeProtokoll_Bargeld_Handling "Sie haben eine Kassendifferenz. Kasse:" & gcKasNum
                schreibeProtokoll_Bargeld_Handling "Winkiss erwartet einen Bargeldbestand von " & Format(ermaktKassensoll, "########0.00") & " Euro. Kasse:" & gcKasNum
                schreibeProtokoll_Bargeld_Handling "M?chten Sie die Differenz von " & Format(dDifferenzbargeld, "########0.00") & " Euro im Kassenbuch vermerken? Kasse:" & gcKasNum
                
                If dDifferenzbargeld > 0 Then
                    sDifftext = "Winkiss erwartet einen Bargeldbestand von " & Format(ermaktKassensoll, "########0.00") & " Euro." & vbCrLf
                    sDifftext = sDifftext & "Sie haben " & Format(Label1(37).Caption, "########0.00") & " Euro gez?hlt." & vbCrLf & vbCrLf
                    sDifftext = sDifftext & "M?chten Sie die Differenz von " & Format(dDifferenzbargeld, "########0.00") & " Euro im Kassenbuch vermerken?" & vbCrLf
                    sDifftext = sDifftext & "Ja = empfohlen"
                ElseIf dDifferenzbargeld < 0 Then
                    sDifftext = "Winkiss erwartet einen Bargeldbestand von " & Format(ermaktKassensoll, "########0.00") & " Euro." & vbCrLf
                    sDifftext = sDifftext & "Sie haben " & Format(Label1(37).Caption, "########0.00") & " Euro gez?hlt." & vbCrLf & vbCrLf
                    sDifftext = sDifftext & "M?chten Sie die Differenz von " & Format(dDifferenzbargeld, "########0.00") & " Euro im Kassenbuch vermerken?" & vbCrLf
                    sDifftext = sDifftext & "Ja = empfohlen"
                End If
                
                iRet = MsgBox(sDifftext, vbQuestion + vbYesNoCancel + vbDefaultButton1, "Winkiss Frage:")
                If iRet = vbYes Then
                    schreibeProtokoll_Bargeld_Handling "Die Frage wurde mit 'Ja' beantwortet. Kasse:" & gcKasNum
                    EinAusKorrektur dDifferenzbargeld
                ElseIf iRet = vbCancel Then
                    schreibeProtokoll_Bargeld_Handling "Die Frage wurde mit 'Nein' beantwortet. Kasse:" & gcKasNum
                    Exit Sub
                End If
            End If
            
            
            
            
            
            'erst Differenz, dann Drucker
            ctmp = "Ist der Drucker funktionsbereit?" & vbCrLf & vbCrLf
            ctmp = ctmp & " Drucker an?" & vbCrLf
            ctmp = ctmp & " Druckerpapier?" & vbCrLf
            iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Winkiss Frage:")
                
            If iRet = vbYes Then
                anzeige "ROT2", "Bitte warten, Belege werden gedruckt...", Label1(49)
            Else
                anzeige "Rot", "Der Vorgang wurde abgebrochen!", Label1(49)
                Exit Sub
                
            End If
            
            
            
            
            
            
            
            
            Screen.MousePointer = 11
        
            If SchreibeDatenWK21b Then
            
                'ohne nachfragen
                If gsZ?hlbeleg = "Listendrucker" Then
                    DruckDatenBargeldWK21b_DINA4 gbQZBON
                Else
                    DruckDatenBargeldWK21b
                End If
                
            End If
            
            gdWechselgeld = 0
            gdKassenGeldGez?hlt = 0
            
            If IsNumeric(Text1(24).Text) Then
                dABSCHOPF = Text1(24).Text
                updateafcstat "ABSCHOPF", dABSCHOPF, gcKasNum
            Else
                dABSCHOPF = 0
            End If
            
            
                
            lKJADate = Fix(Now)
            cKJAZeit = Format$(Now, "HH:MM:SS")
            
            schreibeProtokoll_Bargeld_Handling "gew?hlter Absch?pfungsbetrag: " & Format(dABSCHOPF, "########0.00") & " Euro. Kasse:" & gcKasNum
            
            insertABSCHOPF lKJADate, cKJAZeit, gcKasNum, CInt(gcBedienerNr), dABSCHOPF
            
            If IsNumeric(Text1(25).Text) Then
                lDuBestand = CLng(Text1(25).Text)
               
            Else
                lDuBestand = 0
            End If
            insertDukatenbestand lKJADate, cKJAZeit, gcKasNum, CInt(gcBedienerNr), lDuBestand
            
            If IsNumeric(Label1(48).Caption) Then
                gdWechselgeld = Label1(48).Caption
                
                schreibeProtokoll_Bargeld_Handling "Wechselgeld verbleib in der Kassenschublade: " & Format(gdWechselgeld, "########0.00") & " Euro. Kasse:" & gcKasNum
                
                'sind wir im autolokalen Modus, dann schreiben wir einen neuen Satz in die
                'AFCSTAT von c aleer
                
                 If gbLocalSec Then
                    If gbAutoLokalModus Then
                    
                        Dim sPfad As String
                        Dim lokalDB As DAO.Database
                        Dim sSQL As String
                        Dim lDatum      As Long
                        lDatum = Fix(Now)
                        
                        sPfad = "C:\aLeer"
                        Set lokalDB = OpenDatabase(sPfad & "\kissdata.mdb", False)
                        sSQL = "Insert into AFCSTAT (adate,Kasnum,Wechsel)"
                        sSQL = sSQL & " values ( "
                        sSQL = sSQL & " " & lDatum & " "
                        sSQL = sSQL & ", " & Val(gcKasNum) & " "
                        sSQL = sSQL & ", '" & gdWechselgeld & "' "
                        sSQL = sSQL & " ) "
                        
                        
                        lokalDB.Execute sSQL, dbFailOnError
                        lokalDB.Close
                        
'                        MsgBox gdWechselgeld
                    End If
                End If
                
            End If
            
            '****
            
            If IsNumeric(Label1(37).Caption) Then
                gdKassenGeldGez?hlt = CDbl(Label1(37).Caption)
                gdKassenGeldGez?hlt = gdKassenGeldGez?hlt - dABSCHOPF
            End If
            
           
            

            
            frmWKL21.LeseDatenWKL21
            Unload frmWK21b
            frmWKL21.Show 1
            
        Case 5
        
            Select Case gsEPartner
                Case Is = "ELP"
                    lese_ELPAY_opt
                    setzedrucker gcBonDrucker
                    Kassenschnitt_elPAY
                Case Is = "ZVT"
                    lese_ZVT_opt
                    setzedrucker gcBonDrucker
                    Kassenschnitt_ZVT
                Case Is = "ZV2"
                    lese_ZVT_opt2
                    
                    Kassenschnitt_ZVT2 False
            End Select
            
    End Select
    
    '//new
    If gcWaehrung = "EUR" Then
    
    Else
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 340 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command1_Click"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf."
        
        Fehlermeldung1
        
    End If
End Sub
Private Sub UpdateAFCStatEinUndAuszahlungen(sArt As String, dBetrag As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum          As Long
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    lDatum = Fix(Now)
    
    cSQL = "Select * from AFCSTAT where ADATE = " & Trim$(Str$(lDatum)) & " and KASNUM = " & gcKasNum & " "
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
        
    Select Case sArt
        Case "Auszahlung"
            HoleNeueBonNrWKL20
            If Not IsNull(rsrs!AUSZAHLUNG) Then
                rsrs!AUSZAHLUNG = rsrs!AUSZAHLUNG + dBetrag
            Else
                rsrs!AUSZAHLUNG = dBetrag
            End If
        Case "Einzahlung"
            HoleNeueBonNrWKL20
            If Not IsNull(rsrs!EINZAHLUNG) Then
                rsrs!EINZAHLUNG = rsrs!EINZAHLUNG + dBetrag
            Else
                rsrs!EINZAHLUNG = dBetrag
            End If
    End Select
    'Datum und Kassennummer setzen
    rsrs!ADATE = lDatum
    rsrs!kasnum = Val(gcKasNum)
'    rsrs!BELEGNR = gdBonNr
'    If gdBonNr < CLng(rsrs!BELEGNR) Then
'
'    Else
'        rsrs!BELEGNR = gdBonNr
'    End If
    
    
    
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
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "UpdateAFCStatEinUndAuszahlungen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf."
    
    Fehlermeldung1
End Sub
Private Sub EinAusKorrektur(dDiffbetrag As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim lDatum      As Long
    Dim czeit       As String
    Dim iBedNr      As Integer
    Dim dBetrag     As Double
    Dim cBezeich    As String
    Dim cART        As String
    Dim byKasnum    As Byte
    Dim ctmp        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    If CDbl(Text1(28).Text) = 0 Then
        Exit Sub
    End If
    
    byKasnum = CByte(gcKasNum)
    lDatum = Fix(Now)
    czeit = Format$(Now, "HH:MM:SS")
    
    iBedNr = Val(gcBedienerNr)
    
'    ctmp = Text1(28).Text
'    ctmp = fnMoveComma2Point$(ctmp)
    dBetrag = dDiffbetrag 'Val(ctmp)
 
    If dBetrag < 0 Then
        cBezeich = "KB - Korrektur"
        cART = "AUSZAHLUNG"
        dBetrag = dBetrag * -1
        
'        UpdateAFCStatEinUndAuszahlungen cART, dBetrag
    Else
        cBezeich = "KB - Korrektur"
        cART = "EINZAHLUNG"
'        UpdateAFCStatEinUndAuszahlungen cART, dBetrag
    End If
    
    cSQL = "Select * from KAEINAUS"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    rsrs.AddNew
    rsrs!ADATE = lDatum
    rsrs!AZEIT = czeit
    rsrs!BEDNU = iBedNr
    rsrs!Betrag = dBetrag
    rsrs!BEZEICH = cBezeich
    rsrs!art = cART
    rsrs!kasnum = byKasnum
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from KAEINAUSF"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    rsrs.AddNew
    rsrs!ADATE = lDatum
    rsrs!AZEIT = czeit
    rsrs!BEDNU = iBedNr
    rsrs!Betrag = dBetrag
    rsrs!BEZEICH = cBezeich
    rsrs!art = cART
    rsrs!kasnum = byKasnum
    rsrs!SENDOK = False
    rsrs!FILIALE = gcFilNr
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from EINAUSKB"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    rsrs.AddNew
    rsrs!ADATE = lDatum
    rsrs!AZEIT = czeit
    rsrs!BEDNU = iBedNr
    rsrs!Betrag = dBetrag
    rsrs!BEZEICH = cBezeich
    rsrs!art = cART
    rsrs!kasnum = byKasnum
    rsrs!SENDOK = False
    rsrs!FILIALE = gcFilNr
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    Text1(28).Text = "0"
    Text1(28).Refresh
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EinAusKorrektur"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Vorgang benennen auf. "
    
    Fehlermeldung1
End Sub
Private Sub SendeDaten2DruckerNeuWKL21b()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim iLevel As Integer
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    Dim iFileNr As Integer
    Dim cText As String
    Dim lAnzZeile As Long
    ReDim cDruckZeile(1 To 1) As String
    Dim cDaten As String
    Dim iLenZeile As Integer
    Dim lcount As Long
    Dim ctmp As String
    Dim lAnz As Long
    Dim dWert As Double
    Dim dSumme As Double
    Dim dTotal As Double
    Dim dKassensoll As Double
    Dim dWechselgeld As Double
    Dim dABSCHOPF As Double
    
    iLevel = 0
    
    iLenZeile = 35
    
    '***********************************************
    'Drucker an, Display aus, Init Drucker
    '***********************************************
    
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = Chr$(27) + Chr$(61) + Chr$(1) + Chr$(27) + Chr$(64)
    OpenDrawer aDeviceName, cEscapeSequenz
    
    
    
    '***********************************************
    'ggf. Logo auf Kassenbon bringen
    '***********************************************
    If gcBild <> "" Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '***********************************************
    'Kopfdaten 1.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER F?R DEMO FESTTEXT
        cDaten = "K.I.S.S. Warenwirtschaft"
    Else
        cDaten = gcBonText(0)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    '***********************************************
    'Kopfdaten 2.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER F?R DEMO FESTTEXT
        cDaten = "Carsten Schr?der"
    Else
        cDaten = gcBonText(1)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 3.Zeile an Drucker senden
    '***********************************************
    
    If gbDEMO Then
        'HIER F?R DEMO FESTTEXT
        cDaten = "DEMO-VERSION!"
    Else
        cDaten = gcBonText(4)
    End If
    If Trim$(cDaten) <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Kopfdaten 4.Zeile an Drucker senden
    '***********************************************

    If gbDEMO Then
        'HIER F?R DEMO FESTTEXT
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
    
    '***********************************************
    'Trennstrich drucken
    '***********************************************
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '***********************************************
    'Datenbereich
    '***********************************************
        
    '*** "Kassenabrechnung"
        
    cDaten = "Kassenabrechnung"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    '*** "BARGELD"
        
    cDaten = "BARGELD"
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
        
    '*** Leerzeile
    
    cDaten = Space$(30)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '*** Leerzeile
    
    cDaten = Space$(30)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cSQL = "Select * from DRU_BARG"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        '*** "Filiale"
        
        If Not IsNull(rsrs!FILIALE) Then
            cDaten = rsrs!FILIALE
            cDaten = "Filiale: " & Trim$(cDaten)
        Else
            cDaten = "Filiale: 0"
        End If
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '*** "KASNUM"
        cDaten = "KASNUM: " & gcKasNum
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '*** "Datum"
        
        If Not IsNull(rsrs!Datum) Then
            cDaten = Format$(rsrs!Datum, "DD.MM.YYYY")
            cDaten = "Datum:   " & Trim$(cDaten)
        Else
            cDaten = "Datum:   " & Format$(Now, "DD.MM.YYYY")
        End If
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '*** "Uhrzeit"
        cDaten = "Uhrzeit: " & Format$(TimeValue(Now), "HH:MM:SS")
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '*** "W?hrung"
        
        If Not IsNull(rsrs!Waehrung) Then
            cDaten = rsrs!Waehrung
            cDaten = "W?hrung: " & Trim$(cDaten)
        Else
            cDaten = "W?hrung: EUR"
        End If
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '*** "Wechselgeld vom Kassenstart"
        ctmp = Format(ermaktWechselgeld, "########0.00 ") & gcWaehrung
        ctmp = Trim$(ctmp)
        ctmp = Space$(14 - Len(ctmp)) & ctmp
        cDaten = "Wechselgeld:     " & ctmp
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '*** "Bargeld"
        ctmp = Format(ermaktKassensoll - ermaktWechselgeld, "########0.00 ") & gcWaehrung
        ctmp = Trim$(ctmp)
        ctmp = Space$(18 - Len(ctmp)) & ctmp
        cDaten = "Bargeld:     " & ctmp
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '*** "Anfangsbestand Soll"
        ctmp = Format(ermaktKassensoll, "########0.00 ") & gcWaehrung
        ctmp = Trim$(ctmp)
        ctmp = Space$(15 - Len(ctmp)) & ctmp
        
            
        cDaten = "Kassensoll:     " & ctmp
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        

        
        '*** Leerzeile
        
        cDaten = Space$(30)
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
                
        '*** ?berschrift
        
        cDaten = "Nennwert  Anzahl       Wert"
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '*** Trennstrich
        
        cDaten = String$(iLenZeile, "-")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!NENNWERT) Then
                dWert = rsrs!NENNWERT
            Else
                dWert = 0
            End If
            
            If dWert > 3000 Then
            
            Else
            
            
            
                ctmp = Format$(dWert, "####0.00")
                ctmp = Space$(8 - Len(ctmp)) & ctmp
                cDaten = ctmp & " "
            
                If Not IsNull(rsrs!ANZAHL) Then
                    ctmp = rsrs!ANZAHL
                    lAnz = rsrs!ANZAHL
                Else
                    ctmp = "0"
                    lAnz = 0
                End If
                ctmp = Trim$(ctmp)
                ctmp = Space$(7 - Len(ctmp)) & ctmp
                
                cDaten = cDaten & ctmp & " "
                
                dSumme = lAnz * dWert
                dTotal = dTotal + dSumme
                
                ctmp = Format$(dSumme, "####0.00")
                ctmp = Trim$(ctmp)
                ctmp = Space$(10 - Len(ctmp)) & ctmp
                
                cDaten = cDaten & ctmp
                
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
                
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '*** Trennstrich
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '*** Total
    If dTotal = 0 Then dTotal = CDbl(Label1(37).Caption)
    
    
    
    ctmp = Format$(dTotal, "####0.00")
    ctmp = Trim$(ctmp)
    ctmp = Space$(10 - Len(ctmp)) & ctmp
    
        
    cDaten = "Gesamtsumme:     " & ctmp
    
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    cSQL = "Select * from DRU_BARG where Nennwert = 3007 "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            dABSCHOPF = rsrs!ANZAHL
            ctmp = Format$(dABSCHOPF, "####0.00")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
        Else
            ctmp = "0,00"
        End If
      
        cDaten = "Absch?pfung:     " & ctmp
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '*** Trennstrich
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '*** Wechselgeld
    
    dWechselgeld = 0
    If Label1(48).Caption <> "" Then
        If IsNumeric(Label1(48).Caption) Then
            dWechselgeld = CDbl(Trim(Label1(48).Caption))
        End If
    End If

    ctmp = Format$(dWechselgeld, "####0.00")
    ctmp = Trim$(ctmp)
    ctmp = Space$(10 - Len(ctmp)) & ctmp

    cDaten = "Wechselgeld:     " & ctmp
    
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    
    '*** Leerzeile
    
    cDaten = Space$(30)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz


    'Dukaten
    cSQL = "Select * from DRU_BARG where Nennwert = 3001 "
    cSQL = cSQL & " and Anzahl > 0 and not Anzahl is null "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            ctmp = Format$(rsrs!ANZAHL, "####0.00")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp & " " & gcWaehrung
        Else
            ctmp = "0,00"
        End If
      
        cDaten = "Dukaten:     " & ctmp
        
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'Dukaten
    cSQL = "Select * from DRU_BARG where Nennwert = 3002 "
    cSQL = cSQL & " and Anzahl > 0  and not Anzahl is null  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            ctmp = Format$(rsrs!ANZAHL, "####0")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
        Else
            ctmp = "0,00"
        End If
      
        If ctmp <> "0,00" Then
            cDaten = "Dukaten:     " & ctmp & " St?ck"
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'Dukaten bestand
    cSQL = "Select * from DRU_BARG where Nennwert = 3008 "
    cSQL = cSQL & " and Anzahl > 0  and not Anzahl is null  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            ctmp = Format$(rsrs!ANZAHL, "####0")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
        Else
            ctmp = "0,00"
        End If
      
        If ctmp <> "0,00" Then
            cDaten = "Dukaten B:   " & ctmp & " St?ck"
'            cDaten = "Dukaten:     " & ctmp & " St?ck"
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    'gutscheine
    cSQL = "Select * from DRU_BARG where Nennwert = 3003 "
    cSQL = cSQL & " and Anzahl > 0 and not Anzahl is null "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            ctmp = Format$(rsrs!ANZAHL, "####0.00")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
        Else
            ctmp = "0,00"
        End If
      
        If ctmp <> "0,00" Then
            cDaten = "Gutscheine:  " & ctmp & " " & gcWaehrung
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'gutscheine
    cSQL = "Select * from DRU_BARG where Nennwert = 3004 "
    cSQL = cSQL & " and Anzahl > 0  and not Anzahl is null  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            ctmp = Format$(rsrs!ANZAHL, "####0")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
        Else
            ctmp = "0,00"
        End If
      
        If ctmp <> "0,00" Then
            cDaten = "Gutscheine:  " & ctmp & " St?ck"
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'Kreditkarte
    cSQL = "Select * from DRU_BARG where Nennwert = 3005 "
    cSQL = cSQL & " and Anzahl > 0 and not Anzahl is null "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            ctmp = Format$(rsrs!ANZAHL, "####0.00")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp
        Else
            ctmp = "0,00"
        End If
      
        If ctmp <> "0,00" Then
            cDaten = "Kreditkarte: " & ctmp & " " & gcWaehrung
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'Kreditkarte
    cSQL = "Select * from DRU_BARG where Nennwert = 3006 "
    cSQL = cSQL & " and Anzahl > 0  and not Anzahl is null  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!ANZAHL) Then
            ctmp = Format$(rsrs!ANZAHL, "####0")
            ctmp = Trim$(ctmp)
            ctmp = Space$(10 - Len(ctmp)) & ctmp & " St?ck"
        Else
            ctmp = "0,00"
        End If
      
        If ctmp <> "0,00" Then
            cDaten = "Kreditkarte: " & ctmp
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '***********************************************
    'Ende Datenbereich
    '***********************************************
    
    '*** Leerzeile
    
    cDaten = Space$(30)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '*** Leerzeile
    
    cDaten = Space$(30)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    
    cDaten = "Unterschrift:     "
    
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    '*** Leerzeile
    
    cDaten = Space$(30)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz


    '***********************************************
    'Fu?zeile 1 drucken
    '***********************************************
    
    'Fu?zeilen
    If gbDEMO Then
        'HIER F?R DEMO FESTTEXT
        cDaten = "KEINE G?LTIGE KASSENABRECHNUNG!"
    Else
        cDaten = "ENDE KASSENABRECHNUNG"
    End If
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Fu?zeile 2 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER F?R DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = " "
    End If
    
    
    cDaten = Trim$(cDaten)
    If cDaten <> "" Then
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    iLevel = 10
    
    '***********************************************
    'ein paar Leerzeilen drucken
    '***********************************************
    
    For lcount = 1 To 9
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    Next lcount
    
BON_DRUCKEN:
    
    'OpenDrawer3 benutzt die WindowsAPI
    'OpenDrawer4 geht ?ber das PRINTER-Objekt
    
    If gbAPI = True Then
        OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
    Else
        OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
    End If
   
   
    SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False

BON_SCHNEIDEN:

    'Kassenbon abschneiden
    If gbAPI = True Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = Chr$(27) + Chr$(105)
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    iLevel = 11
    
    Erase cDruckZeile
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerNeuWKL21b"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. " & Trim$(Str$(iLevel))
    
    Fehlermeldung1
    Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Dim iCount As Integer
    
    Screen.MousePointer = 11
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    Label1(4).Visible = False
    Text1(2).Visible = False
    Label1(20).Visible = False
        
    Label1(37).Caption = "0,00"
    
    For iCount = 0 To 27
        Text1(iCount).Text = ""
    Next iCount
    
    For iCount = 20 To 35
        Label1(iCount).Caption = "0,00"
    Next iCount
    
    If gbBarAnz = True Then
         Label1(49).Caption = "Dieser Betrag wird erwartet: " & Format(ermaktKassensoll, "########0.00") & " " & gcWaehrung
         Label1(49).Refresh
    End If
    
    If gbBarAnz = False Then
         Text1(28).Visible = False
         Label1(52).Visible = False
    End If
    
    Text1(1).Text = Format$(Now, "DD.MM.YYYY")
    Text1(26).Text = gcWaehrung
    
    Text1(0).Text = gcFilNr
    
    If gbBARZSCHUB Then
        'schublade ?ffnen
        SchubladeOeffnen
    End If
    
    If gbBargeldEingabe = True Then
        Command1(4).Visible = True
        
        Command1(0).Visible = True
        Command1(1).Visible = False
        Command1(2).Visible = True
        Command1(3).Visible = False
        
    Else
        Command1(4).Visible = False
        
        Command1(0).Visible = True
        Command1(1).Visible = True
        Command1(2).Visible = True
        Command1(3).Visible = True
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_Change(index As Integer)
On Error GoTo LOKAL_ERROR

    Dim dSum As Double
    Dim dABSCHOPF As Double
    Dim dWert As Double
    Dim i As Integer
    
    dWert = 0
    
    If index = 24 Then
        If IsNumeric(Label1(37).Caption) Then
            dSum = CDbl(Label1(37).Caption)
            If IsNumeric(Text1(24).Text) Then
                dABSCHOPF = CDbl(Text1(24).Text)
            Else
                dABSCHOPF = 0
            End If
            
            dWert = dSum - dABSCHOPF
            
            If dWert < 0 Then
                Text1(24).Text = ""
                Text1(24).SetFocus
                anzeige "rot", "", Label1(48)
                anzeige "normal", "", Label1(48)
            Else
                Label1(48).Caption = Format$(dWert, "###,##0.00")
            End If
        Else
            Label1(48).Caption = Format$(dWert, "###,##0.00")
        End If
    ElseIf index >= 2 And index <= 17 Then
    
        Label1(37).Caption = Format$(Zaehlezusammen, "###,##0.00")
        Label1(37).Refresh
        
        If Text1(24).Text <> "" Then
            If IsNumeric(Text1(24).Text) Then
                Text1_Change 24
            End If
        End If
    
        If IsNumeric(Text1(index).Text) Then
            Text1(27).Text = ""
        End If
        
        Text1(28).Text = Format$(ermaktKassensoll - CDbl(Label1(37).Caption), "###,##0.00")
    
    ElseIf index = 27 Then
        If IsNumeric(Text1(27).Text) Then
        
            For i = 2 To 17
                Text1(i).Text = ""
            Next i
            
            Label1(37).Caption = Format$(Text1(27).Text, "###,##0.00")
            Label1(37).Refresh
            
            If Text1(24).Text <> "" Then
                If IsNumeric(Text1(24).Text) Then
                    Text1_Change 24
                End If
            End If
            
            Text1(28).Text = Format$(ermaktKassensoll - CDbl(Label1(37).Caption), "###,##0.00")
        End If
    End If

    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
End Sub
Private Function Zaehlezusammen() As Double
On Error GoTo LOKAL_ERROR

Dim dSum As Double
Dim i As Integer

Zaehlezusammen = 0
dSum = 0

For i = 2 To 17
    If IsNumeric(Text1(i).Text) Then
        dSum = dSum + (Val(Text1(i).Text) * CDbl(Label1(i + 2).Tag))
        
        Label1(i + 18).Caption = Format(Val(Text1(i).Text) * CDbl(Label1(i + 2).Tag), "###,##0.00")
    End If
Next i

Zaehlezusammen = dSum

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zaehlezusammen"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
End Function
Private Sub Text1_GotFocus(index As Integer)
On Error GoTo LOKAL_ERROR

    Label0.Caption = Trim$(Str$(index))
    Text1(index).BackColor = glSelBack1
    Text1(index).SelStart = 0
    Text1(index).SelLength = Len(Text1(index).Text)
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    cZeichen = Chr$(KeyAscii)
    
    Select Case index
    
        Case Is = 18, 20, 22, 24, 27
            cValid = "1234567890," & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
            If cZeichen = "," Then
                If InStr(Text1(index).Text, ",") > 0 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
    
        Case Else
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
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

    If index >= 2 And index <= 18 Then
    
        If KeyCode = vbKeyReturn Then
            Text1(index + 1).SetFocus
        End If
        
    ElseIf index = 19 Then
        If KeyCode = vbKeyReturn Then
            Text1(25).SetFocus
        End If
    ElseIf index = 25 Then
        If KeyCode = vbKeyReturn Then
            Text1(20).SetFocus
        End If
    ElseIf index = 20 Then
        If KeyCode = vbKeyReturn Then
            Text1(index + 1).SetFocus
        End If
    ElseIf index = 21 Then
        If KeyCode = vbKeyReturn Then
            Text1(index + 1).SetFocus
        End If
    ElseIf index = 22 Then
        If KeyCode = vbKeyReturn Then
            Text1(index + 1).SetFocus
        End If
    ElseIf index = 23 Then
        If KeyCode = vbKeyReturn Then
            Text1(index + 1).SetFocus
        End If
    ElseIf index = 27 Then
        If KeyCode = vbKeyReturn Then
            Text1(24).SetFocus
        End If
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_LostFocus(index As Integer)
On Error GoTo LOKAL_ERROR

    Text1(index).BackColor = vbWhite

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kasseninhalt auf. "
    
    Fehlermeldung1
End Sub


