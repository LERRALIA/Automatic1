VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL101 
   Caption         =   "Kunde neu"
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
      Index           =   9
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "12"
      Text            =   "Text1"
      Top             =   4080
      Width           =   2175
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
      Index           =   8
      Left            =   120
      MaxLength       =   13
      TabIndex        =   7
      Tag             =   "15"
      Text            =   "14"
      Top             =   3360
      Width           =   2055
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
      Index           =   7
      Left            =   9120
      MaxLength       =   35
      TabIndex        =   14
      Tag             =   "14"
      Text            =   "13"
      Top             =   4080
      Width           =   2655
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
      Index           =   6
      Left            =   7440
      MaxLength       =   35
      TabIndex        =   13
      Tag             =   "13"
      Text            =   "12"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CheckBox Checkbox1 
      Caption         =   "mit Groß/Klein Automatik"
      Height          =   375
      Left            =   7680
      TabIndex        =   42
      Top             =   240
      Width           =   3375
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
      Index           =   5
      Left            =   2880
      MaxLength       =   35
      TabIndex        =   5
      Tag             =   "6"
      Text            =   "Text1"
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   975
      Left            =   360
      TabIndex        =   37
      Top             =   8160
      Visible         =   0   'False
      Width           =   3255
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   4
         Left            =   5280
         TabIndex        =   41
         Top             =   2760
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
         Caption         =   "Änder / Ausw"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   2
         Left            =   7440
         TabIndex        =   40
         Top             =   2760
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
         Caption         =   "Ignorieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   1
         Left            =   9600
         TabIndex        =   39
         Top             =   2760
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
         Caption         =   "Auswählen"
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
         Height          =   1950
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   11535
      End
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   0
      Left            =   7920
      TabIndex        =   15
      Top             =   7920
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   33
      Top             =   4560
      Visible         =   0   'False
      Width           =   11775
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   47
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   12
         Left            =   1005
         TabIndex        =   48
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   13
         Left            =   1770
         TabIndex        =   49
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   14
         Left            =   2535
         TabIndex        =   50
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   15
         Left            =   3300
         TabIndex        =   51
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   16
         Left            =   4065
         TabIndex        =   52
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   17
         Left            =   4830
         TabIndex        =   53
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   18
         Left            =   5595
         TabIndex        =   54
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   19
         Left            =   6360
         TabIndex        =   55
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   20
         Left            =   7125
         TabIndex        =   56
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   21
         Left            =   7890
         TabIndex        =   57
         Top             =   600
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "ß"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   22
         Left            =   360
         TabIndex        =   58
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Q"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   23
         Left            =   1120
         TabIndex        =   59
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "W"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   24
         Left            =   1890
         TabIndex        =   60
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "E"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   25
         Left            =   2655
         TabIndex        =   61
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "R"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   26
         Left            =   3420
         TabIndex        =   62
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "T"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   27
         Left            =   4185
         TabIndex        =   63
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Z"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   28
         Left            =   4950
         TabIndex        =   64
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "U"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   29
         Left            =   5715
         TabIndex        =   65
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "I"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   30
         Left            =   6480
         TabIndex        =   66
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "O"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   31
         Left            =   7245
         TabIndex        =   67
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "P"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   32
         Left            =   8010
         TabIndex        =   68
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Ü"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   34
         Left            =   8780
         TabIndex        =   69
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   1
         Left            =   9540
         TabIndex        =   70
         Top             =   1240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "@"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   35
         Left            =   480
         TabIndex        =   71
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   36
         Left            =   1245
         TabIndex        =   72
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "S"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   37
         Left            =   2010
         TabIndex        =   73
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "D"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   38
         Left            =   2775
         TabIndex        =   74
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "F"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   39
         Left            =   3540
         TabIndex        =   75
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "G"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   40
         Left            =   4305
         TabIndex        =   76
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "H"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   41
         Left            =   5070
         TabIndex        =   77
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "J"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   42
         Left            =   5835
         TabIndex        =   78
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "K"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   43
         Left            =   6600
         TabIndex        =   79
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "L"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   44
         Left            =   7365
         TabIndex        =   80
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Ö"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   45
         Left            =   8130
         TabIndex        =   81
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Ä"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   46
         Left            =   8900
         TabIndex        =   82
         Top             =   1890
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "/"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   47
         Left            =   600
         TabIndex        =   83
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Y"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   48
         Left            =   1365
         TabIndex        =   84
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "X"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   49
         Left            =   2130
         TabIndex        =   85
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   50
         Left            =   2895
         TabIndex        =   86
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "V"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   51
         Left            =   3660
         TabIndex        =   87
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "B"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   52
         Left            =   4425
         TabIndex        =   88
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "N"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   53
         Left            =   5190
         TabIndex        =   89
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "M"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   54
         Left            =   5955
         TabIndex        =   90
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ";"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   55
         Left            =   6720
         TabIndex        =   91
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ":"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   56
         Left            =   7485
         TabIndex        =   92
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "_"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   57
         Left            =   8250
         TabIndex        =   93
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   " "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   58
         Left            =   9015
         TabIndex        =   94
         Top             =   2540
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   0
         Left            =   9660
         TabIndex        =   95
         Top             =   1890
         Width           =   1820
         _Version        =   65536
         _ExtentX        =   3210
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Leeren"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   1
         Left            =   9780
         TabIndex        =   96
         Top             =   2540
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Rückg"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   2
         Left            =   8660
         TabIndex        =   97
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "<<<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   3
         Left            =   9900
         TabIndex        =   98
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   ">>>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Index           =   4
         Left            =   10300
         TabIndex        =   99
         Top             =   1240
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "A -> a"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "-1"
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
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         Caption         =   "Zielfeld:"
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
         Height          =   255
         Index           =   0
         Left            =   8640
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
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
      Index           =   4
      Left            =   2880
      MaxLength       =   20
      TabIndex        =   11
      Tag             =   "11"
      Text            =   "Text1"
      Top             =   4080
      Width           =   2175
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
      Index           =   3
      Left            =   120
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "10"
      Text            =   "Text1"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   6
      Tag             =   "7"
      Text            =   "Combo1"
      Top             =   2640
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   8
      Tag             =   "8"
      Text            =   "Combo1"
      Top             =   3360
      Width           =   2055
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
      Index           =   2
      Left            =   7440
      MaxLength       =   35
      TabIndex        =   4
      Tag             =   "5"
      Text            =   "Text1"
      Top             =   1920
      Width           =   4335
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
      Index           =   1
      Left            =   2880
      MaxLength       =   35
      TabIndex        =   3
      Tag             =   "4"
      Text            =   "Text1"
      Top             =   1920
      Width           =   4335
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
      Index           =   0
      Left            =   7440
      MaxLength       =   35
      TabIndex        =   2
      Tag             =   "3"
      Text            =   "Text1"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   7440
      TabIndex        =   9
      Tag             =   "9"
      Text            =   "Combo1"
      Top             =   3360
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Tag             =   "2"
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Tag             =   "1"
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1815
   End
   Begin sevCommand3.Command Command6 
      Height          =   375
      Index           =   11
      Left            =   11400
      TabIndex        =   19
      Top             =   240
      Width           =   375
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
      Caption         =   "?"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   3
      Left            =   9840
      TabIndex        =   16
      Top             =   7920
      Width           =   1815
      _ExtentX        =   3201
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
   Begin sevCommand3.Command Command0 
      Height          =   360
      Index           =   1
      Left            =   2280
      TabIndex        =   46
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
   Begin sevCommand3.Command Command1 
      Height          =   495
      Index           =   5
      Left            =   6000
      TabIndex        =   101
      Top             =   7920
      Width           =   1815
      _ExtentX        =   3201
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
      Index           =   6
      Left            =   5160
      TabIndex        =   102
      Top             =   7920
      Width           =   735
      _ExtentX        =   1296
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
      Caption         =   "DS"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Mobil"
      Height          =   255
      Index           =   15
      Left            =   5160
      TabIndex        =   100
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Kundenkarte"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   45
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Email"
      Height          =   255
      Index           =   13
      Left            =   9120
      TabIndex        =   44
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Merkmal"
      Height          =   255
      Index           =   12
      Left            =   7440
      TabIndex        =   43
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Land"
      Height          =   255
      Index           =   11
      Left            =   7440
      TabIndex        =   32
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Firma"
      Height          =   255
      Index           =   10
      Left            =   7440
      TabIndex        =   31
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Telefon"
      Height          =   255
      Index           =   9
      Left            =   2880
      TabIndex        =   30
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Geburtsdatum"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   29
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Strasse"
      Height          =   255
      Index           =   7
      Left            =   2880
      TabIndex        =   28
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Ort"
      Height          =   255
      Index           =   6
      Left            =   7440
      TabIndex        =   27
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Plz"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   26
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nachname"
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   25
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Vorname"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   24
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Titel"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   22
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Anrede"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   21
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Kundennummer"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "Kunde neu"
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
      TabIndex        =   18
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lbl1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   8040
      Width           =   4935
   End
End
Attribute VB_Name = "frmWKL101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bfil As Boolean



Private Sub Combo1_GotFocus(Index As Integer)
On Error GoTo LOKAL_ERROR

    Label3(2).Caption = Combo1(Index).Tag
    
    Combo1(Index).BackColor = glSelBack1
    Combo1(Index).SelStart = 0
    Combo1(Index).SelLength = Len(Combo1(Index).Text)
    
    Select Case Index
        Case 4 'stadt
            If bfil = True Then
                fülleSpaltemitKrit Combo1(4), "Stadtd", "Kustadt", "Stadtd", Combo1(4).Text, "", " where plz like '" & Combo1(3).Text & "*'"
            End If
        Case 3 'plz
            If bfil = True Then
                If Text1(5).Text <> "" Then
                    fülleSpaltemitKrit Combo1(3), "PLZ", "Kunden", "PLZ", Combo1(3).Text, "", " where stadt like '" & Combo1(4).Text & "*'   and strasse like '" & Left(Text1(5).Text, 6) & "*'"
                Else
                    fülleSpaltemitKrit Combo1(3), "PLZd", "Kuplz", "PLZd", Combo1(3).Text, "", " where stadt like '" & Combo1(4).Text & "*'"
                End If
            End If
    End Select
    
    bfil = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Combo1_LostFocus(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Combo1(Index).BackColor = vbWhite
    Select Case Index
        Case 3
            If Trim(Combo1(3).Text) <> "" Then
                If Len(Trim(Combo1(3).Text)) < 5 Then
                    If UCase(Trim(Combo1(2).Text)) = "DEUTSCHLAND" Then
                        Combo1(2).Text = ""
                    End If
                Else
                    If Len(Trim(Combo1(3).Text)) = 5 Then
                        If Trim(Combo1(2).Text) = "" Then
                            Combo1(2).Text = "Deutschland"
                        End If
                    End If
                End If
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Combo1_LostFocus"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Select Case Index
        Case Is = 1        ' Kalender
            Text1(3).Text = Format(Datumschreiben11a(3000, 4000), "DD.MM.YYYY")
            'fertig
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cKundnr As String
    
    Select Case Index
        Case 0
            If pruef(True) Then
                SpeicherNeuKunde True
            End If
        Case 1
            gcKundenNr = ""
            If List1.ListCount > 0 Then
                For lcount = 0 To List1.ListCount - 1
                    If List1.Selected(lcount) = True Then
                        gcKundenNr = Trim(Left(List1.list(lcount), 10))
                        Exit For
                    End If
                Next lcount
                
                If gcKundenNr <> "" Then
                    Unload Me
                End If
            End If
        Case 2
            If pruef(False) Then
                SpeicherNeuKunde True
            End If
        Case 3
            Unload frmWKL101
        Case 4
            gcKundenNr = ""
            If List1.ListCount > 0 Then
                For lcount = 0 To List1.ListCount - 1
                    If List1.Selected(lcount) = True Then
                        gcKundenNr = Trim(Left(List1.list(lcount), 10))
                        Exit For
                    End If
                Next lcount
                
                If gcKundenNr <> "" Then
                    iKasse = 2
                    frmWKL13.Show 1
'                    setzedrucker gcBonDrucker
                    Command1_Click 1
                End If
            End If
            
        Case 5
        
            If pruef(True) Then
                SpeicherNeuKunde False
            End If
        
            cKundnr = Label2.Caption
            cKundnr = Trim$(cKundnr)
            
            If cKundnr = "" Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            
            setzedrucker gcListenDrucker

            StammdatenblattKundeDrucken cKundnr, False
                
            setzedrucker gcBonDrucker
                
                
            
            'auf Bondrucker
            'StammdatenblattKundeDrucken_Bondrucker cKundnr
            
        Case 6
        
            If pruef(True) Then
                SpeicherNeuKunde False
            End If
        
            cKundnr = Label2.Caption
            cKundnr = Trim$(cKundnr)
            
            If cKundnr = "" Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            
            setzedrucker gcListenDrucker

            DatenschutzblattKundeDrucken cKundnr
                
            setzedrucker gcBonDrucker

        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SSCommand1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    iZielIndex = Label3(2).Caption

    Select Case iZielIndex
        Case 1
            Combo1(0).BackColor = glSelBack1
        Case 2
            Combo1(1).BackColor = glSelBack1
        Case 3
            Text1(0).BackColor = glSelBack1
        Case 4
            Text1(1).BackColor = glSelBack1
        Case 5
            Text1(2).BackColor = glSelBack1
        Case 6
            Text1(5).BackColor = glSelBack1
        Case 7
            Combo1(4).BackColor = glSelBack1
        Case 8
            Combo1(3).BackColor = glSelBack1
        Case 9
            Combo1(2).BackColor = glSelBack1
        Case 10
            Text1(3).BackColor = glSelBack1
        Case 11
            Text1(4).BackColor = glSelBack1
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub
Private Sub schalte(bWie As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Combo1(0).Enabled = bWie
    Combo1(1).Enabled = bWie
    Text1(0).Enabled = bWie
    Text1(1).Enabled = bWie
    Text1(2).Enabled = bWie
    Combo1(3).Enabled = bWie
    Combo1(4).Enabled = bWie
    Text1(5).Enabled = bWie
    Combo1(2).Enabled = bWie
    Text1(3).Enabled = bWie
    Text1(4).Enabled = bWie
    Command1(0).Visible = bWie
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "schalte"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
    
End Sub
Private Sub SSCommand2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim iZielIndex As Integer
    Dim lcount As Long
    
    If Label3(2).Caption <> "" Then
        iZielIndex = Label3(2).Caption
         Select Case Index
             Case Is = 0     'CLEAR
                 Select Case iZielIndex
                     Case 1
                         Combo1(0).Text = ""
                         Combo1(0).SetFocus
                     Case 2
                         Combo1(1).Text = ""
                         Combo1(1).SetFocus
                     Case 3
                         Text1(0).Text = ""
                         Text1(0).SetFocus
                     Case 4
                         Text1(1).Text = ""
                         Text1(1).SetFocus
                     Case 5
                         Text1(2).Text = ""
                         Text1(2).SetFocus
                    Case 6
                         Text1(5).Text = ""
                         Text1(5).SetFocus
                    Case 7
                        bfil = False
                        fülleSpaltemitKrit Combo1(4), "Stadtd", "Kustadt", "Stadtd", "", "", ""
                        Combo1(4).SetFocus
                     Case 8
                        bfil = False
                        fülleSpaltemitKrit Combo1(3), "PLZd", "Kuplz", "PLZd", "", "", ""
                        Combo1(3).SetFocus
                    Case 9
                        Combo1(2).Text = ""
                         Combo1(2).SetFocus
                    Case 10
                         Text1(3).Text = ""
                         Text1(3).SetFocus
                    Case 11
                         Text1(4).Text = ""
                         Text1(4).SetFocus
                    Case 12
                         Text1(9).Text = ""
                         Text1(9).SetFocus
                    Case 13
                         Text1(6).Text = ""
                         Text1(6).SetFocus
                    Case 14
                         Text1(7).Text = ""
                         Text1(7).SetFocus
                 End Select
                
            Case Is = 1     'rück
                Select Case iZielIndex
                     Case 1
                        If Len(Combo1(0).Text) > 0 Then
                            Combo1(0).Text = Left(Combo1(0).Text, Len(Combo1(0).Text) - 1)
                        End If
                        Combo1(0).SetFocus
                     Case 2
                        If Len(Combo1(1).Text) > 0 Then
                            Combo1(1).Text = Left(Combo1(1).Text, Len(Combo1(1).Text) - 1)
                        End If
                        Combo1(1).SetFocus
                     Case 3
                        If Len(Text1(0).Text) > 0 Then
                            Text1(0).Text = Left(Text1(0).Text, Len(Text1(0).Text) - 1)
                        End If
                        Text1(0).SetFocus
                     Case 4
                        If Len(Text1(1).Text) > 0 Then
                            Text1(1).Text = Left(Text1(1).Text, Len(Text1(1).Text) - 1)
                        End If
                        Text1(1).SetFocus
                     Case 5
                        If Len(Text1(2).Text) > 0 Then
                            Text1(2).Text = Left(Text1(2).Text, Len(Text1(2).Text) - 1)
                        End If
                        Text1(2).SetFocus
                     Case 6
                        If Len(Text1(5).Text) > 0 Then
                            Text1(5).Text = Left(Text1(5).Text, Len(Text1(5).Text) - 1)
                        End If
                        Text1(5).SetFocus
                        
                     Case 7
                        If Len(Combo1(4).Text) > 0 Then
                            Combo1(4).Text = Left(Combo1(4).Text, Len(Combo1(4).Text) - 1)
                        End If
                        Combo1(4).SetFocus
                     Case 8
                        If Len(Combo1(3).Text) > 0 Then
                            Combo1(3).Text = Left(Combo1(3).Text, Len(Combo1(3).Text) - 1)
                        End If
                        Combo1(3).SetFocus
                        
                     Case 9
                        If Len(Combo1(2).Text) > 0 Then
                            Combo1(2).Text = Left(Combo1(2).Text, Len(Combo1(2).Text) - 1)
                        End If
                        Combo1(2).SetFocus
                     Case 10
                        If Len(Text1(3).Text) > 0 Then
                            Text1(3).Text = Left(Text1(3).Text, Len(Text1(3).Text) - 1)
                        End If
                        Text1(3).SetFocus
                     Case 11
                        If Len(Text1(4).Text) > 0 Then
                            Text1(4).Text = Left(Text1(4).Text, Len(Text1(4).Text) - 1)
                        End If
                        Text1(4).SetFocus
                    Case 12
                        If Len(Text1(9).Text) > 0 Then
                            Text1(9).Text = Left(Text1(9).Text, Len(Text1(9).Text) - 1)
                        End If
                        Text1(9).SetFocus
                     Case 13
                        If Len(Text1(6).Text) > 0 Then
                            Text1(6).Text = Left(Text1(6).Text, Len(Text1(6).Text) - 1)
                        End If
                        Text1(6).SetFocus
                    Case 14
                        If Len(Text1(7).Text) > 0 Then
                            Text1(7).Text = Left(Text1(7).Text, Len(Text1(7).Text) - 1)
                        End If
                        Text1(7).SetFocus
                 End Select
                 
 
         
            Case Is = 2     'BEFORE
            

                Select Case iZielIndex
                    Case 1
                        Combo1(0).SetFocus
                    Case 2
                         Combo1(0).SetFocus
                    Case 3
                         Combo1(1).SetFocus
                    Case 4
                         Text1(0).SetFocus
                    Case 5
                         Text1(1).SetFocus
                    Case 6
                         Text1(2).SetFocus
                    Case 7
                         Text1(5).SetFocus
                    Case 8
                        Text1(8).SetFocus
                    Case 9
                        Combo1(3).SetFocus

                    Case 10
                         Combo1(2).SetFocus
                    Case 11
                         Text1(3).SetFocus
                    Case 12
                         Text1(4).SetFocus
                         
                    Case 13
                         
                         Text1(9).SetFocus
                         
                    Case 14
                         Text1(6).SetFocus
                    Case 15
                        Combo1(4).SetFocus
                 End Select
             
            Case Is = 3     'NEXT

                Select Case iZielIndex
                    Case 1
                        Combo1(1).SetFocus
                    Case 2
                        Text1(0).SetFocus
                    Case 3
                        Text1(1).SetFocus
                    Case 4
                        Text1(2).SetFocus
                    Case 5
                        Text1(5).SetFocus
                    Case 6
                        Combo1(4).SetFocus
                    Case 7
                        Text1(8).SetFocus
                    Case 8
                        Combo1(2).SetFocus
                    Case 9
                        Text1(3).SetFocus
                    Case 10
                        Text1(4).SetFocus
                    Case 11
                        Text1(9).SetFocus
                        
                    Case 12
                        Text1(6).SetFocus
                    Case 13
                        Text1(7).SetFocus
                    Case 14
                        Text1(7).SetFocus
                    Case 15
                        Combo1(3).SetFocus
                 End Select
                    
                    
            Case Is = 4     'switch UPPER CASE / lower case
                SwitchUpperLowerCaseWKL13
        End Select
        
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command6_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11

    Select Case Index
        Case 11 'Hilfe
            gsHelpstring = "Kasse \ Kunde neu"
            frmWKL110.Show 1
    End Select
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command6_Click"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim zIndex As Integer
    
  
    If Label3(2).Caption <> "" Then
        
        zIndex = CInt(Label3(2).Caption)
        
        Select Case zIndex
            Case 1
                Combo1(0).Text = Combo1(0).Text & SSCommand1(Index).Caption
                Combo1(0).SetFocus
            Case 2
                Combo1(1).Text = Combo1(1).Text & SSCommand1(Index).Caption
                Combo1(1).SetFocus
            Case 3
                
                Text1(0).Text = Text1(0).Text & SSCommand1(Index).Caption
                Text1(0).SetFocus
            Case 4
                Text1(1).Text = Text1(1).Text & SSCommand1(Index).Caption
                Text1(1).SetFocus
            Case 5
                Text1(2).Text = Text1(2).Text & SSCommand1(Index).Caption
                Text1(2).SetFocus
            Case 6
                Text1(5).Text = Text1(5).Text & SSCommand1(Index).Caption
                Text1(5).SetFocus
                
            Case 7
                Combo1(4).Text = Combo1(4).Text & SSCommand1(Index).Caption
                Combo1(4).SetFocus
            Case 8
                If IsNumeric(SSCommand1(Index).Caption) Then
                    Combo1(3).Text = Combo1(3).Text & SSCommand1(Index).Caption
                End If
                Combo1(3).SetFocus
                
            Case 9
                Combo1(2).Text = Combo1(2).Text & SSCommand1(Index).Caption
                Combo1(2).SetFocus
            Case 10
                Text1(3).Text = Text1(3).Text & SSCommand1(Index).Caption
                Text1(3).SetFocus
            Case 11
                If IsNumeric(SSCommand1(Index).Caption) Then
                    Text1(4).Text = Text1(4).Text & SSCommand1(Index).Caption
                End If
                Text1(4).SetFocus
            Case 12
                Text1(9).Text = Text1(9).Text & SSCommand1(Index).Caption
                Text1(9).SetFocus
            Case 13
                Text1(6).Text = Text1(6).Text & SSCommand1(Index).Caption
                Text1(6).SetFocus
            Case 14
                Text1(7).Text = Text1(7).Text & SSCommand1(Index).Caption
                Text1(7).SetFocus
        End Select
       
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    WKL101Positionieren
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    alternativFarbform Me, lblUeberschrift

    LogtoStart Me
    bfil = True
    BereiteVor
    
    If NewTableSuchenDBKombi("E101", gdApp) Then
        voreinstellungladen101
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungladen101()
    On Error GoTo LOKAL_ERROR

    Dim rs As Recordset

    Set rs = gdApp.OpenRecordset("E101")
    If Not rs.EOF Then
        If rs!bo7 = True Then
            Checkbox1.Value = vbUnchecked
        Else
            Checkbox1.Value = vbChecked
        End If
    End If
    rs.Close: Set rs = Nothing


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungladen101"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub voreinstellungspeichern101()
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    Dim bo7 As Integer

    loeschNEW "E101", gdApp
    CreateTable "E101", gdApp

    
    If Checkbox1.Value = vbChecked Then
        bo7 = 0
    Else
        bo7 = -1
    End If
    
    sSQL = "Insert into E101 ( bo7) "
    sSQL = sSQL & " values (" & bo7
    sSQL = sSQL & " )"
    gdApp.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "voreinstellungspeichern101"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub

Private Sub WKL101Positionieren()
    On Error GoTo LOKAL_ERROR
    
    Frame2.Top = 4560
    Frame2.Left = 0
    Frame2.Height = 3255
    Frame2.Width = 11655
    
    Frame1.Top = 4560
    Frame1.Left = 0
    Frame1.Height = 3255
    Frame1.Width = 11775
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WKL101Positionieren"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SwitchUpperLowerCaseWKL13()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    
    If Left(SSCommand2(4).Caption, 1) = "A" Then
        For lcount = 22 To 32
            SSCommand1(lcount).Caption = LCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 35 To 45
            SSCommand1(lcount).Caption = LCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 47 To 55
            SSCommand1(lcount).Caption = LCase(SSCommand1(lcount).Caption)
        Next lcount
        SSCommand1(54).Caption = ","
        SSCommand1(55).Caption = "."
        SSCommand1(56).Caption = "-"
        
        SSCommand2(4).Caption = "a -> A"
    Else
        For lcount = 22 To 32
            SSCommand1(lcount).Caption = UCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 35 To 45
            SSCommand1(lcount).Caption = UCase(SSCommand1(lcount).Caption)
        Next lcount
        For lcount = 47 To 55
            SSCommand1(lcount).Caption = UCase(SSCommand1(lcount).Caption)
        Next lcount
        
        SSCommand1(54).Caption = ";"
        SSCommand1(55).Caption = ":"
        SSCommand1(56).Caption = "_"
        
        SSCommand2(4).Caption = "A -> a"
    
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SwitchUpperLowerCaseWKL13"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo LOKAL_ERROR
    
    voreinstellungspeichern101
    LogtoEnd Me
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Unload"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cZeichen As String
    Dim cValid As String
    
    
    Select Case Index
        Case 0, 1, 2, 5 'Firma Name Vorname Strasse
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß%"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            
        Case 3 'Gebdat
            cValid = "1234567890." & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 4, 8 'Telefon
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
    Fehler.gsFehlertext = "Im Programmteil Kundendaten bearbeiten ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Label3(2).Caption = Text1(Index).Tag

    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BereiteVor()
    On Error GoTo LOKAL_ERROR
    Dim i As Integer
    
    Label2.Caption = ""
    Label2.Caption = Format$(fnHoleMaxKundenNr, "#####0")
    
    fülleSpalte Combo1(1), "Titeld", "KUTITEL", "Titeld", "", ""
    fuellecombo
    
    For i = 0 To 9
        Text1(i).Text = ""
    Next i
    
    fülleSpalte Combo1(3), "PLZd", "Kuplz", "PLZd", "", ""
    fülleSpalte Combo1(4), "Stadtd", "Kustadt", "Stadtd", "", ""
    
    If gbBILDTAST = False Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BereiteVor"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub SpeicherNeuKunde(bVerlassen As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim ctmp As String
    Dim dWert As Double
    
    sSQL = "Select * from KUNDEN where KUNDNR = " & Label2.Caption & " "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
    
    Else
    
        anzeige "normal", "Daten werden gespeichert...", lbl1
        rsrs.AddNew
        rsrs!Kundnr = Label2.Caption
        rsrs!Kuerzel = UCase(Left(Text1(2).Text, 5))
        
        Select Case Trim$(Combo1(0).Text)
            Case Is = "Frau"
                rsrs!geschlecht = "W"
            Case Is = "Herr"
                rsrs!geschlecht = "M"
            Case Is = "Familie"
                rsrs!geschlecht = "F"
            Case Is = "Firma"
                rsrs!geschlecht = "U"
        End Select
        
        rsrs!firma = Trim$(Text1(0).Text)
        rsrs!titel = Left(Trim$(Combo1(1).Text), 35)
        rsrs!name = Trim$(Text1(2).Text)
        rsrs!vorname = Trim$(Text1(1).Text)
        rsrs!strasse = Trim$(Text1(5).Text)
        rsrs!Plz = Left(Trim$(Combo1(3).Text), 7)
        rsrs!STADT = Left(Trim$(Combo1(4).Text), 35)
        rsrs!Tel = Trim$(Text1(4).Text)
        rsrs!Mobiltel = Trim$(Text1(9).Text)
        rsrs!FAXNR = ""
        rsrs!anrede = Left(Trim$(Combo1(0).Text), 10)
        rsrs!Email = Trim$(Text1(7).Text)
        rsrs!KUNDKART = Trim$(Text1(8).Text)
        rsrs!MERKMAL = Trim$(Text1(6).Text)
        rsrs!MERKMAL2 = "J"
        rsrs!FORMATDAT = ""
        rsrs!Rechnr = gcBedienerNr
        
        If Combo1(2).Text <> "" Then
            Select Case Combo1(2).Text
                Case "Deutschland"
                    rsrs!KURZTEXT1 = "D"
                Case "Schweiz"
                    rsrs!KURZTEXT1 = "CH"
                Case "Österreich"
                    rsrs!KURZTEXT1 = "A"
                Case "Belgien"
                    rsrs!KURZTEXT1 = "B"
                Case "Dänemark"
                    rsrs!KURZTEXT1 = "DK"
                Case "Frankreich"
                    rsrs!KURZTEXT1 = "F"
                Case "Italien"
                    rsrs!KURZTEXT1 = "I"
                Case "Lichtenstein"
                    rsrs!KURZTEXT1 = "FL"
                Case "Luxemburg"
                    rsrs!KURZTEXT1 = "L"
                Case "Monaco"
                    rsrs!KURZTEXT1 = "Mo"
                Case "Niederlande"
                    rsrs!KURZTEXT1 = "NL"
                Case "Polen"
                    rsrs!KURZTEXT1 = "PL"
                Case "Portugal"
                    rsrs!KURZTEXT1 = "P"
                Case "Spanien"
                    rsrs!KURZTEXT1 = "E"
                Case Else
                    rsrs!KURZTEXT1 = ""
            End Select
        End If
        
        rsrs!KurzTEXT2 = ""
        rsrs!AWM = "0"
        rsrs!RABATT = 0
        
        ctmp = Text1(3).Text
        If ctmp <> "" Then
            dWert = DateValue(ctmp)
            rsrs!Datum1 = dWert
            If gbGEBRABK Then
                If Month(rsrs!Datum1) = Month(DateValue(Now)) Then
                    rsrs!AWM = "5"
                    rsrs!RABATT = 10
                End If
            End If
        Else
            rsrs!Datum1 = Null
        End If
        
        rsrs!DATUM2 = Null
        rsrs!UMSLJ = 0
        rsrs!UMSVJ = 0
        rsrs!OSUM = 0
        
        rsrs!Status = "A"
        rsrs!SYNStatus = "A"
        rsrs!TBONUS = 0
        rsrs!BONUS = 0
        rsrs!ECIDENT = ""
        rsrs!GESPERRT = "N"
        rsrs!NOTIZEN = ""
        
        sSQL = "Delete from  BANKKU where kundnr = " & Trim$(Label2.Caption)
        gdBase.Execute sSQL, dbFailOnError
        
        rsrs!DS = False
        rsrs!BE = False
        rsrs!PREISKZ = 0
        rsrs!FILIALNR = Val(gcFilNr)
        rsrs!angelegt = Fix(Now)
        rsrs!AENDER = "I"
        rsrs!LASTDATE = DateValue(Now)
        rsrs!LASTTIME = TimeValue(Now)
        rsrs.Update
        
        anzeige "normal", "", lbl1
        
        If bVerlassen = True Then
            gcKundenNr = Label2.Caption
            frmWKL20.Text3(3).Text = gcKundenNr
            Unload Me
        End If

    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SpeicherNeuKunde"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function pruef(bmitDuplicheck As Boolean) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    pruef = False
    
    'Pflicht
    If Trim(Text1(2).Text) = "" Then
        anzeige "rot", "Der Name fehlt", lbl1
        Text1(2).SetFocus
        Exit Function
    Else
        If Checkbox1.Value = vbChecked Then
            Text1(2).Text = GKAutomatik(Text1(2).Text, "NACHNAME")
        End If
    End If
    
    If Trim(Combo1(0).Text) = "" Then
        anzeige "rot", "Die Anrede fehlt", lbl1
        Combo1(0).SetFocus
        Exit Function
    Else
    
        'anrede
        If Trim(Combo1(0).Text) = "" Then
            
        Else
            If Checkbox1.Value = vbChecked Then
                Combo1(0).Text = GKAutomatik(Combo1(0).Text, "ANREDE")
                
            End If
        End If
        
        Select Case Trim$(Combo1(0).Text)
            Case Is = "Frau"
                
            Case Is = "Herr"
                
            Case Is = "Familie"
                
            Case Is = "Firma"
            
            Case Else
            
            anzeige "rot", "Die Anrede ist falsch", lbl1
            Combo1(0).SetFocus
            Exit Function
        End Select
    End If
    
    'land
    If Trim(Combo1(2).Text) = "" Then
        
    Else
        If Checkbox1.Value = vbChecked Then
            Combo1(2).Text = GKAutomatik(Combo1(2).Text, "LAND")
        End If
    End If
    
    If Trim(Combo1(2).Text) = "" Then
        anzeige "rot", "Das Land fehlt", lbl1
        Combo1(2).SetFocus
        Exit Function
    End If
    
    If Trim(Combo1(3).Text) <> "" Then
        If Len(Combo1(3).Text) < 5 And Combo1(2).Text = "Deutschland" Then

            anzeige "rot", "Plz ist zu klein oder Land falsch", lbl1
            Combo1(3).SetFocus
            Exit Function

        End If
    End If
    
    If Trim(Text1(3).Text) <> "" Then
        If IsDate(Text1(3).Text) Then
        
        Else
            anzeige "rot", "Das Datum ist falsch", lbl1
            Text1(3).SetFocus
            Exit Function
        
        End If
    End If
    
    'Email
    If Trim(Text1(7).Text) <> "" Then
        If InStr(Text1(7).Text, "@") = 0 Then
            anzeige "rot", "Die Emailadresse ist falsch. Es fehlt ein '@'.", lbl1
            Text1(7).SetFocus
            Exit Function
        
        End If
    End If
    
    'Email
    If Trim(Text1(7).Text) <> "" Then
        If InStr(Text1(7).Text, ".") = 0 Then
            anzeige "rot", "Die Emailadresse ist falsch. Es fehlt ein Punkt.", lbl1
            Text1(7).SetFocus
            Exit Function
        
        End If
    End If
    
    'Vorname
    If Trim(Text1(1).Text) = "" Then
        
    Else
        If Checkbox1.Value = vbChecked Then
            Text1(1).Text = GKAutomatik(Text1(1).Text, "VORNAME")
        End If
    End If
    
    'firma
    If Trim(Text1(0).Text) = "" Then
        
    Else
        If Checkbox1.Value = vbChecked Then
            Text1(0).Text = GKAutomatik(Text1(0).Text, "FIRMA")
        End If
    End If
    
    'strasse
    If Trim(Text1(5).Text) = "" Then
        
    Else
        If Checkbox1.Value = vbChecked Then
            Text1(5).Text = GKAutomatik(Text1(5).Text, "STRASSE")
        End If
    End If
    
    'ort
    If Trim(Combo1(4).Text) = "" Then
        
    Else
        If Checkbox1.Value = vbChecked Then
            Combo1(4).Text = GKAutomatik(Combo1(4).Text, "ORT")
        End If
    End If
    
    'titel
    If Trim(Combo1(1).Text) = "" Then
        
    Else
        If Checkbox1.Value = vbChecked Then
            Combo1(1).Text = GKAutomatik(Combo1(1).Text, "TITEL")
        End If
    End If
    
    
    'und jetzt duplicheck
    If bmitDuplicheck Then
    
        If duplicheck = False Then
            
            anzeige "rot", "schon angelegt?", lbl1
            
            schalte False
            Frame2.Visible = False
            Frame1.Visible = True
            
            Exit Function
        Else
        
            schalte True
            Frame2.Visible = True
            Frame1.Visible = False
        
        End If
    Else
    
        schalte True
        Frame2.Visible = True
        Frame1.Visible = False
        
    End If
    
    'Ende
    
    
    
    pruef = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Pruef"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function GKAutomatik(sChecktext As String, sArt As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim lStart As Long
    Dim cFeld As String
    Dim cEinzelteil As String
    Dim cGesamt As String
    Dim lPos As Long
    Dim cSeek As String
    
    Select Case sArt
    
        Case "VORNAME"
        
            lStart = 1
            cEinzelteil = ""
            cGesamt = ""
            cFeld = Trim(sChecktext)
            cFeld = cFeld & " "
            
            Do While lStart < Len(cFeld)
                lPos = InStr(lStart, cFeld, " ")
                If lPos <> 0 Then
                    cEinzelteil = Mid(cFeld, lStart, lPos - lStart)
                    If cEinzelteil = "" Then
                    
                    Else
                        cGesamt = cGesamt & Left(UCase(Trim(cEinzelteil)), 1) & Right(LCase(Trim(cEinzelteil)), Len(Trim(cEinzelteil)) - 1)
                        cGesamt = cGesamt & " "
                    End If
                End If

                lStart = lPos + 1
                If lStart = 0 Then
                    Exit Do
                End If
                
            Loop
        
            GKAutomatik = Trim(cGesamt)
        
        Case "NACHNAME"
        
            lStart = 1
            cEinzelteil = ""
            cGesamt = ""
            cFeld = Trim(sChecktext)
            cFeld = cFeld & " "
            
            Do While lStart < Len(cFeld)
                lPos = InStr(lStart, cFeld, " ")
                If lPos <> 0 Then
                    cEinzelteil = Mid(cFeld, lStart, lPos - lStart)
                    If cEinzelteil = "" Then
                    
                    ElseIf UCase(cEinzelteil) = "VON" Then
                        cGesamt = cGesamt & "von "
                    ElseIf UCase(cEinzelteil) = "VAN" Then
                        cGesamt = cGesamt & "van "
                    ElseIf UCase(cEinzelteil) = "DI" Then
                        cGesamt = cGesamt & "di "
                    ElseIf UCase(cEinzelteil) = "DER" Then
                        cGesamt = cGesamt & "der "
                    Else
                        cGesamt = cGesamt & Left(UCase(Trim(cEinzelteil)), 1) & Right(LCase(Trim(cEinzelteil)), Len(Trim(cEinzelteil)) - 1)
                        cGesamt = cGesamt & " "
                    End If
                End If

                lStart = lPos + 1
                If lStart = 0 Then
                    Exit Do
                End If
                
            Loop
        
            GKAutomatik = Trim(cGesamt)
        Case "ANREDE"
            GKAutomatik = Left(UCase(Trim(sChecktext)), 1) & Right(LCase(Trim(sChecktext)), Len(Trim(sChecktext)) - 1)
        Case "LAND"
            GKAutomatik = Left(UCase(Trim(sChecktext)), 1) & Right(LCase(Trim(sChecktext)), Len(Trim(sChecktext)) - 1)
        Case "FIRMA"
            GKAutomatik = Left(UCase(Trim(sChecktext)), 1) & Right(LCase(Trim(sChecktext)), Len(Trim(sChecktext)) - 1)
        Case "ORT"
            GKAutomatik = Left(UCase(Trim(sChecktext)), 1) & Right(LCase(Trim(sChecktext)), Len(Trim(sChecktext)) - 1)
        Case "TITEL"
            GKAutomatik = Left(UCase(Trim(sChecktext)), 1) & Right(LCase(Trim(sChecktext)), Len(Trim(sChecktext)) - 1)
        Case "STRASSE"
            lStart = 1
            cEinzelteil = ""
            cGesamt = ""
            
            
            If InStr(lStart, sChecktext, "-") > 0 Then
                cSeek = "-"
                sChecktext = SwapStr(sChecktext, " ", "-")
            Else
                cSeek = " "
                sChecktext = SwapStr(sChecktext, "-", " ")
            End If
            
            cFeld = Trim(sChecktext)
            cFeld = cFeld & cSeek
            
        
            Do While lStart < Len(cFeld)
                lPos = InStr(lStart, cFeld, cSeek)
                If lPos <> 0 Then
                    cEinzelteil = Mid(cFeld, lStart, lPos - lStart)
                    If cEinzelteil = "" Then
                    
                    ElseIf UCase(cEinzelteil) = "VON" Then
                        cGesamt = cGesamt & "von " & cSeek
                    ElseIf UCase(cEinzelteil) = "VAN" Then
                        cGesamt = cGesamt & "van" & cSeek
                    ElseIf UCase(cEinzelteil) = "DI" Then
                        cGesamt = cGesamt & "di" & cSeek
                    ElseIf UCase(cEinzelteil) = "DER" Then
                        cGesamt = cGesamt & "der" & cSeek
                    Else
                        cGesamt = cGesamt & Left(UCase(Trim(cEinzelteil)), 1) & Right(LCase(Trim(cEinzelteil)), Len(Trim(cEinzelteil)) - 1)
                        cGesamt = cGesamt & cSeek
                    End If
                End If

                lStart = lPos + 1
                If lStart = 0 Then
                    Exit Do
                End If
                
            Loop
            
            If cSeek = " " Then
        
                GKAutomatik = Trim(cGesamt)
            ElseIf cSeek = "-" Then
                GKAutomatik = Left(cGesamt, Len(cGesamt) - 1)
            End If
    
    End Select

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GKAutomatik"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function duplicheck() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim cSatz   As String
    Dim ctmp    As String

    
    duplicheck = True
    List1.Clear
    

    Screen.MousePointer = 11
    
    anzeige "normal", "Duplikate werden gesucht...", lbl1
    
    If Text1(3).Text <> "" Then
        sSQL = "Select "
        sSQL = sSQL & " name  "
        sSQL = sSQL & " ,vorname  "
        sSQL = sSQL & " ,titel  "
        sSQL = sSQL & " ,strasse  "
        sSQL = sSQL & " ,stadt  "
        sSQL = sSQL & " ,plz  "
        sSQL = sSQL & " ,tel  "
        sSQL = sSQL & " ,datum1  "
        sSQL = sSQL & " ,kundnr  "
        sSQL = sSQL & " from Kunden where "
        sSQL = sSQL & " ucase(name) = '" & UCase(Trim(Text1(2).Text)) & "'"
        sSQL = sSQL & " and ucase(vorname) = '" & UCase(Trim(Text1(1).Text)) & "'"
        sSQL = sSQL & " and datum1 = " & CLng(DateValue(Text1(3).Text)) & " "
'        sSQL = sSQL & " and Filialnr = " & gcFilNr

        
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            duplicheck = False
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                cSatz = ""

                If Not IsNull(rsrs!Kundnr) Then
                ctmp = rsrs!Kundnr
            End If
            cSatz = cSatz & Space$(10 - Len(ctmp)) & ctmp


            If Not IsNull(rsrs!titel) Then
                ctmp = rsrs!titel
                ctmp = Left(ctmp, 8)
            Else
                ctmp = Space(8)
            End If
            cSatz = cSatz & Space$(9 - Len(ctmp)) & Left(ctmp, 8)

            If Not IsNull(rsrs!name) Then
                ctmp = rsrs!name
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If
            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!vorname) Then
                ctmp = rsrs!vorname
                ctmp = Left(ctmp, 14)
            Else
                ctmp = Space(14)
            End If
            cSatz = cSatz & Space$(15 - Len(ctmp)) & Left(ctmp, 14)

            If Not IsNull(rsrs!Plz) Then
                ctmp = rsrs!Plz
            Else
                ctmp = Space(7)
            End If
            cSatz = cSatz & Space$(8 - Len(ctmp)) & Left(ctmp, 7)

            If Not IsNull(rsrs!STADT) Then
                ctmp = rsrs!STADT
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If

            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!strasse) Then
                ctmp = rsrs!strasse
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If
            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!Datum1) Then
                ctmp = rsrs!Datum1
            Else
                ctmp = Space(10)
            End If
            cSatz = cSatz & Space$(11 - Len(ctmp)) & Left(ctmp, 10) & Space(2)

            If Not IsNull(rsrs!Tel) Then
                ctmp = rsrs!Tel
                ctmp = Left(ctmp, 20)
            Else
                ctmp = Space(20)
            End If
            cSatz = cSatz & Left(ctmp, 20)

                List1.AddItem cSatz

            rsrs.MoveNext
            Loop

        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    sSQL = " Select "
    sSQL = sSQL & " name  "
    sSQL = sSQL & " ,vorname  "
    sSQL = sSQL & " ,titel  "
    sSQL = sSQL & " ,strasse  "
    sSQL = sSQL & " ,stadt  "
    sSQL = sSQL & " ,plz  "
    sSQL = sSQL & " ,tel  "
    sSQL = sSQL & " ,datum1  "
    sSQL = sSQL & " ,kundnr  "
    sSQL = sSQL & " from Kunden where "
    sSQL = sSQL & " ucase(name) = '" & UCase(Trim(Text1(2).Text)) & "'"
    sSQL = sSQL & " and ucase(left(strasse,6)) = '" & UCase(Left(Trim(Text1(5).Text), 6)) & "'"
    sSQL = sSQL & " and ucase(vorname) = '" & UCase(Trim(Text1(1).Text)) & "'"

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        duplicheck = False
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cSatz = ""

            If Not IsNull(rsrs!Kundnr) Then
                ctmp = rsrs!Kundnr
            End If
            cSatz = cSatz & Space$(10 - Len(ctmp)) & ctmp


            If Not IsNull(rsrs!titel) Then
                ctmp = rsrs!titel
                ctmp = Left(ctmp, 8)
            Else
                ctmp = Space(8)
            End If
            cSatz = cSatz & Space$(9 - Len(ctmp)) & Left(ctmp, 8)

            If Not IsNull(rsrs!name) Then
                ctmp = rsrs!name
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If
            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!vorname) Then
                ctmp = rsrs!vorname
                ctmp = Left(ctmp, 14)
            Else
                ctmp = Space(14)
            End If
            cSatz = cSatz & Space$(15 - Len(ctmp)) & Left(ctmp, 14)

            If Not IsNull(rsrs!Plz) Then
                ctmp = rsrs!Plz
            Else
                ctmp = Space(7)
            End If
            cSatz = cSatz & Space$(8 - Len(ctmp)) & Left(ctmp, 7)

            If Not IsNull(rsrs!STADT) Then
                ctmp = rsrs!STADT
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If

            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!strasse) Then
                ctmp = rsrs!strasse
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If
            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!Datum1) Then
                ctmp = rsrs!Datum1
            Else
                ctmp = Space(10)
            End If
            cSatz = cSatz & Space$(11 - Len(ctmp)) & Left(ctmp, 10) & Space(2)

            If Not IsNull(rsrs!Tel) Then
                ctmp = rsrs!Tel
                ctmp = Left(ctmp, 20)
            Else
                ctmp = Space(20)
            End If
            cSatz = cSatz & Left(ctmp, 20)

            List1.AddItem cSatz

        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing

    sSQL = "Select "
    sSQL = sSQL & " name  "
    sSQL = sSQL & " ,vorname  "
    sSQL = sSQL & " ,titel  "
    sSQL = sSQL & " ,strasse  "
    sSQL = sSQL & " ,stadt  "
    sSQL = sSQL & " ,plz  "
    sSQL = sSQL & " ,tel  "
    sSQL = sSQL & " ,datum1  "
    sSQL = sSQL & " ,kundnr  "
    sSQL = sSQL & " from Kunden where "
    sSQL = sSQL & " ucase(name) = '" & UCase(Trim(Text1(2).Text)) & "'"
    sSQL = sSQL & " and ucase(vorname) = '" & UCase(Trim(Text1(1).Text)) & "'"
    sSQL = sSQL & " and ucase(stadt) = '" & UCase(Trim(Combo1(4).Text)) & "'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        duplicheck = False
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cSatz = ""

            If Not IsNull(rsrs!Kundnr) Then
                ctmp = rsrs!Kundnr
            End If
            cSatz = cSatz & Space$(10 - Len(ctmp)) & ctmp


            If Not IsNull(rsrs!titel) Then
                ctmp = rsrs!titel
                ctmp = Left(ctmp, 8)
            Else
                ctmp = Space(8)
            End If
            cSatz = cSatz & Space$(9 - Len(ctmp)) & Left(ctmp, 8)

            If Not IsNull(rsrs!name) Then
                ctmp = rsrs!name
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If
            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!vorname) Then
                ctmp = rsrs!vorname
                ctmp = Left(ctmp, 14)
            Else
                ctmp = Space(14)
            End If
            cSatz = cSatz & Space$(15 - Len(ctmp)) & Left(ctmp, 14)

            If Not IsNull(rsrs!Plz) Then
                ctmp = rsrs!Plz
            Else
                ctmp = Space(7)
            End If
            cSatz = cSatz & Space$(8 - Len(ctmp)) & Left(ctmp, 7)

            If Not IsNull(rsrs!STADT) Then
                ctmp = rsrs!STADT
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If

            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!strasse) Then
                ctmp = rsrs!strasse
                ctmp = Left(ctmp, 18)
            Else
                ctmp = Space(18)
            End If
            cSatz = cSatz & Space$(20 - Len(ctmp)) & Left(ctmp, 18)

            If Not IsNull(rsrs!Datum1) Then
                ctmp = rsrs!Datum1
            Else
                ctmp = Space(10)
            End If
            cSatz = cSatz & Space$(11 - Len(ctmp)) & Left(ctmp, 10) & Space(2)

            If Not IsNull(rsrs!Tel) Then
                ctmp = rsrs!Tel
                ctmp = Left(ctmp, 20)
            Else
                ctmp = Space(20)
            End If
            cSatz = cSatz & Left(ctmp, 20)

            List1.AddItem cSatz

        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing

    Screen.MousePointer = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "duplicheck"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub fuellecombo()
    On Error GoTo LOKAL_ERROR
    
    
    Combo1(0).Clear
    Combo1(0).AddItem "Frau"
    Combo1(0).AddItem "Herr"
    Combo1(0).AddItem "Familie"
    Combo1(0).AddItem "Firma"
    Combo1(0).Text = ""
    
    Combo1(2).Clear
    Combo1(2).AddItem "Deutschland"
    Combo1(2).AddItem "Schweiz"
    Combo1(2).AddItem "Österreich"
    
    Combo1(2).AddItem "Belgien"
    Combo1(2).AddItem "Dänemark"
    Combo1(2).AddItem "Frankreich"
    Combo1(2).AddItem "Italien"
    Combo1(2).AddItem "Lichtenstein"
    Combo1(2).AddItem "Luxemburg"
    Combo1(2).AddItem "Monaco"
    Combo1(2).AddItem "Niederlande"
    Combo1(2).AddItem "Polen"
    Combo1(2).AddItem "Portugal"
    Combo1(2).AddItem "Spanien"
    
    Combo1(2).Text = "Deutschland"
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fuellecombo"
    Fehler.gsFehlertext = "Im Programmteil Kunde neu ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub


