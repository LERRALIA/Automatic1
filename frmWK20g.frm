VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWK20g 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "WinKISS - Zahlung per Gutschein"
   ClientHeight    =   8625
   ClientLeft      =   1830
   ClientTop       =   1515
   ClientWidth     =   11910
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.Frame Frame10 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2895
      Left            =   0
      TabIndex        =   98
      Top             =   1320
      Visible         =   0   'False
      Width           =   6495
      Begin sevCommand3.Command Command5 
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   100
         Top             =   2400
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   1508
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
         Caption         =   "OK"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   855
         Index           =   1
         Left            =   3720
         TabIndex        =   99
         Top             =   2400
         Width           =   3255
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
      Begin VB.Label Label28 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1935
         Left            =   240
         TabIndex        =   105
         Top             =   3720
         Width           =   6735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         Caption         =   "Summe:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   104
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00808000&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   5
         Left            =   2280
         TabIndex        =   103
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00808000&
         Caption         =   "LEER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   240
         TabIndex        =   102
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00808000&
         Caption         =   "€"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   6360
         TabIndex        =   101
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00FF0000&
      Height          =   4695
      Left            =   5640
      TabIndex        =   52
      Top             =   2760
      Visible         =   0   'False
      Width           =   7215
      Begin MSCommLib.MSComm MSComm1 
         Left            =   120
         Top             =   720
         _ExtentX        =   794
         _ExtentY        =   794
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   69
         Text            =   "Text3"
         Top             =   1080
         Width           =   1575
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   11
         Left            =   4560
         TabIndex        =   82
         Top             =   2520
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   10
         Left            =   3840
         TabIndex        =   81
         Top             =   2520
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   9
         Left            =   3120
         TabIndex        =   80
         Top             =   2520
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   8
         Left            =   2400
         TabIndex        =   79
         Top             =   2520
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   7
         Left            =   1680
         TabIndex        =   78
         Top             =   2520
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   6
         Left            =   960
         TabIndex        =   77
         Top             =   2520
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   5
         Left            =   240
         TabIndex        =   76
         Top             =   2520
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   4
         Left            =   3120
         TabIndex        =   75
         Top             =   1800
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   3
         Left            =   2400
         TabIndex        =   74
         Top             =   1800
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   2
         Left            =   1680
         TabIndex        =   73
         Top             =   1800
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   1
         Left            =   960
         TabIndex        =   72
         Top             =   1800
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin Threed.SSCommand SSCommand7 
         Height          =   700
         Index           =   0
         Left            =   240
         TabIndex        =   71
         Top             =   1800
         Width           =   700
         _Version        =   65536
         _ExtentX        =   1235
         _ExtentY        =   1235
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
      Begin sevCommand3.Command SSCommand8 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   113
         Top             =   3360
         Width           =   1670
         _ExtentX        =   2937
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
         Caption         =   "Bar"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   615
         Index           =   1
         Left            =   1920
         TabIndex        =   114
         Top             =   3360
         Width           =   1670
         _ExtentX        =   2937
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
         Caption         =   "Scheck"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   615
         Index           =   2
         Left            =   3600
         TabIndex        =   115
         Top             =   3360
         Width           =   1670
         _ExtentX        =   2937
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
         Caption         =   "Karte"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   116
         Top             =   3990
         Width           =   1670
         _ExtentX        =   2937
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
         Caption         =   "EC-Last"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand8 
         Height          =   615
         Index           =   4
         Left            =   3600
         TabIndex        =   117
         Top             =   3990
         Width           =   1670
         _ExtentX        =   2937
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         Caption         =   "noch zu zahlen"
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
         Index           =   3
         Left            =   120
         TabIndex        =   96
         Top             =   120
         Width           =   3495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         Index           =   0
         X1              =   240
         X2              =   5880
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   10
         Left            =   4680
         TabIndex        =   70
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "gegeben:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   9
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   8
         Left            =   4680
         TabIndex        =   67
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   2880
         TabIndex        =   66
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "offener Betrag:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   7
         Left            =   240
         TabIndex        =   65
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2040
      TabIndex        =   40
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox Text5 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         MaxLength       =   9
         TabIndex        =   48
         Text            =   "0,00"
         Top             =   840
         Width           =   1695
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   12
         Left            =   3720
         TabIndex        =   83
         Top             =   1920
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
         _StockProps     =   78
         Caption         =   "-"
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   11
         Left            =   4440
         TabIndex        =   64
         Top             =   2640
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   10
         Left            =   3720
         TabIndex        =   63
         Top             =   2640
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   9
         Left            =   3000
         TabIndex        =   62
         Top             =   2640
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   8
         Left            =   2280
         TabIndex        =   61
         Top             =   2640
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   7
         Left            =   1560
         TabIndex        =   60
         Top             =   2640
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   6
         Left            =   840
         TabIndex        =   59
         Top             =   2640
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   5
         Left            =   120
         TabIndex        =   58
         Top             =   2640
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   4
         Left            =   3000
         TabIndex        =   57
         Top             =   1920
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   3
         Left            =   2280
         TabIndex        =   56
         Top             =   1920
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   2
         Left            =   1560
         TabIndex        =   55
         Top             =   1920
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   1
         Left            =   840
         TabIndex        =   54
         Top             =   1920
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand6 
         Height          =   705
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   1920
         Width           =   705
         _Version        =   65536
         _ExtentX        =   1244
         _ExtentY        =   1244
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
      Begin Threed.SSCommand SSCommand5 
         Height          =   615
         Index           =   1
         Left            =   2880
         TabIndex        =   51
         Top             =   3480
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "Abbrechen"
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
      Begin Threed.SSCommand SSCommand5 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   3480
         Width           =   2740
         _Version        =   65536
         _ExtentX        =   4851
         _ExtentY        =   1085
         _StockProps     =   78
         Caption         =   "OK"
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
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         Caption         =   "Rückgeld"
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
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   6
         Left            =   4920
         TabIndex        =   49
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   5
         Left            =   4920
         TabIndex        =   47
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   4
         Left            =   4920
         TabIndex        =   46
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   1
         Left            =   3120
         TabIndex        =   45
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   3120
         TabIndex        =   44
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Rest-Gutschein:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   43
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "davon in bar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   42
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Rückgeld:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   41
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
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
      Height          =   855
      Left            =   8280
      TabIndex        =   23
      Top             =   5160
      Width           =   3495
      Begin VB.TextBox Text2 
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
         Left            =   2280
         MaxLength       =   7
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   720
         Width           =   2415
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   13
         Left            =   2520
         TabIndex        =   39
         Top             =   3120
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Abbrechen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   12
         Left            =   360
         TabIndex        =   38
         Top             =   3120
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   1296
         _StockProps     =   78
         Caption         =   "Speichern"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   11
         Left            =   3960
         TabIndex        =   37
         Top             =   2280
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   10
         Left            =   3960
         TabIndex        =   36
         Top             =   1560
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   9
         Left            =   3240
         TabIndex        =   35
         Top             =   2280
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   8
         Left            =   2520
         TabIndex        =   34
         Top             =   2280
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   7
         Left            =   1800
         TabIndex        =   33
         Top             =   2280
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   6
         Left            =   1080
         TabIndex        =   32
         Top             =   2280
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   5
         Left            =   360
         TabIndex        =   31
         Top             =   2280
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   4
         Left            =   3240
         TabIndex        =   30
         Top             =   1560
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   3
         Left            =   2520
         TabIndex        =   29
         Top             =   1560
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   2
         Left            =   1800
         TabIndex        =   28
         Top             =   1560
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   1
         Left            =   1080
         TabIndex        =   27
         Top             =   1560
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   735
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   1296
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
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         Caption         =   "Alten Gutschein nacherfassen"
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
         Index           =   4
         Left            =   120
         TabIndex        =   97
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Wert:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   4335
      Left            =   5040
      TabIndex        =   21
      Top             =   480
      Width           =   5535
      Begin VB.TextBox Text4 
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
         Height          =   1935
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   106
         Top             =   2400
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   5415
      End
      Begin sevCommand3.Command SSCommand3 
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   111
         Top             =   1680
         Width           =   2730
         _ExtentX        =   4815
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
         Caption         =   "zurücknehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand3 
         Height          =   615
         Index           =   1
         Left            =   2760
         TabIndex        =   112
         Top             =   1680
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "löschen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "ausgewählte Gutscheine"
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
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   84
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
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
      Height          =   2055
      Left            =   360
      TabIndex        =   20
      Top             =   120
      Width           =   5295
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "zu zahlen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   94
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   0
         Left            =   2400
         TabIndex        =   93
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   3
         Left            =   4440
         TabIndex        =   92
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   4
         Left            =   4440
         TabIndex        =   91
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   90
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Gutschein(e):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   89
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   5
         Left            =   4440
         TabIndex        =   88
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   2400
         TabIndex        =   87
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "noch offen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   86
         Top             =   1320
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   7560
      Width           =   11895
      Begin sevCommand3.Command SSCommand2 
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   107
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1296
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
         Caption         =   "Gutschein auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand2 
         Height          =   735
         Index           =   1
         Left            =   3280
         TabIndex        =   108
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
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
         Caption         =   "alter Gutschein"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand2 
         Height          =   735
         Index           =   2
         Left            =   5970
         TabIndex        =   109
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
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
         Caption         =   "Abbrechen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command SSCommand2 
         Height          =   735
         Index           =   3
         Left            =   8660
         TabIndex        =   110
         Top             =   0
         Width           =   2900
         _ExtentX        =   5106
         _ExtentY        =   1296
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
         Caption         =   "Kassieren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   6615
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4020
         Left            =   240
         TabIndex        =   18
         Top             =   3000
         Width           =   5775
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   2640
         Width           =   5775
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
         Height          =   420
         Left            =   2400
         MaxLength       =   13
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   13
         Left            =   4560
         TabIndex        =   16
         Top             =   1800
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "C"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   12
         Left            =   3960
         TabIndex        =   15
         Top             =   1800
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   ","
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   11
         Left            =   3360
         TabIndex        =   14
         Top             =   1800
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   10
         Left            =   2760
         TabIndex        =   13
         Top             =   1800
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   9
         Left            =   2160
         TabIndex        =   12
         Top             =   1800
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   8
         Left            =   1560
         TabIndex        =   11
         Top             =   1800
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   7
         Left            =   960
         TabIndex        =   10
         Top             =   1800
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   6
         Left            =   3360
         TabIndex        =   9
         Top             =   1200
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   5
         Left            =   2760
         TabIndex        =   8
         Top             =   1200
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   4
         Left            =   2160
         TabIndex        =   7
         Top             =   1200
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   3
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   580
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Top             =   1200
         Width           =   580
         _Version        =   65536
         _ExtentX        =   1023
         _ExtentY        =   1023
         _StockProps     =   78
         Caption         =   "1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   1
         Left            =   5040
         TabIndex        =   1
         Top             =   720
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Nr."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Index           =   0
         Left            =   4080
         TabIndex        =   2
         Top             =   720
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Wert"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderWidth     =   2
         Index           =   1
         X1              =   6120
         X2              =   6120
         Y1              =   120
         Y2              =   7080
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         Caption         =   "vorhandene Gutscheine"
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
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   85
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche Gutschein:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmWK20g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim dZuZahlen               As Double
Dim dEingereichteGutscheine As Double
Dim dNochOffen              As Double
Dim dWertRestGutschein      As Double
Dim dGegeben                As Double
Dim bseekerfolg             As Boolean
Dim gBAgeschlossen          As Boolean
Dim sSort                   As String
Dim bFirstEingabe           As Boolean
Public Sub Kassieren()
    On Error GoTo LOKAL_ERROR
    
    TSSBerechnung
    SendeDaten2DruckerGutschein2WK20g 0, dZuZahlen, dGegeben, dWertRestGutschein
    UpdateAFCStatGutscheinModul20 dZuZahlen, dEingereichteGutscheine, dNochOffen, dWertRestGutschein
    InsertAFCBuchGutscheinModul20 dEingereichteGutscheine
    If CheckofP = True Then
        InsertProvision
    End If
    If CheckofX = True Then
        InsertXMarkierung
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Kassieren"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
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
Private Sub WK20gPositionieren()
    On Error GoTo LOKAL_ERROR

    Frame1.Top = 120
    Frame1.Left = 0
    Frame1.Height = 7335
    Frame1.Width = 6255
    Frame1.Visible = True
    
    Frame2.Top = 7560
    Frame2.Left = 120
    Frame2.Height = 855
    Frame2.Width = 11895
    Frame2.Visible = True
    
    Frame3.Top = 240
    Frame3.Left = 6240
    Frame3.Height = 1935
    Frame3.Width = 5415
    Frame3.Visible = True
    
    Frame5.Top = 2400
    Frame5.Left = 6240
    Frame5.Height = 4455 '2295
    Frame5.Width = 5415
    Frame5.Visible = True
    
    Frame6.Height = 4215
    Frame6.Width = 5055
    Frame6.Top = 2280
    Frame6.Left = 3960
    Frame6.Visible = False
    
    Frame7.Height = 4215
    Frame7.Width = 5775
    Frame7.Top = 2040
    Frame7.Left = 3120
    Frame7.Visible = False
    
    Frame8.Height = 4935
    Frame8.Width = 5535
    Frame8.Top = 1440
    Frame8.Left = 3120
    Frame8.Visible = False
    
    Frame10.Height = 5415
    Frame10.Width = 7455
    Frame10.Top = 1320
    Frame10.Left = 2400
    Frame10.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WK20gPositionieren"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
End Sub
Private Sub AktualisiereZahlungWK20g(cLBSatz As String, cOperation As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lHeute          As Long
    Dim cGutschnr       As String
    Dim cSQL            As String
    Dim cFeld           As String
    
    Dim dNeuWert        As Double
    Dim dAltWert        As Double
    Dim cWert           As String
    
    
    lHeute = Fix(Now)
    cGutschnr = Left(cLBSatz, 8)
    cGutschnr = Trim$(cGutschnr)
    

    
    'Wert des Gutscheines ermitteln
    cFeld = Mid(cLBSatz, 22, 9)
    cFeld = fnMoveComma2Point$(cFeld)
    dNeuWert = Val(cFeld)
    
    'zu zahlenden Betrag ermitteln
    cFeld = Label3(0).Caption
    cFeld = fnMoveComma2Point$(cFeld)
    dZuZahlen = Val(cFeld)
    
    'bereits durch Gutscheine gezahlten Betrag ermitteln
    cFeld = Label3(1).Caption
    cFeld = fnMoveComma2Point$(cFeld)
    dAltWert = Val(cFeld)
    
    'neuen Wert für bereits durch Gutscheine gezahlten Betrag ermitteln
    If cOperation = "+" Then
        dAltWert = dAltWert + dNeuWert
        
        cSQL = "Update GUTSCH set DAT_EINL = " & Trim$(Str$(lHeute)) & " where GUTSCHNR = " & cGutschnr
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update GUTSCH set SYNStatus = 'E' where GUTSCHNR = " & cGutschnr
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
        
        cSQL = "Update GUTSCH set Status = 'E' where GUTSCHNR = " & cGutschnr
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

    End If
    
    If cOperation = "-" Then
        dAltWert = dAltWert - dNeuWert
        cSQL = "Update GUTSCH set DAT_EINL = NULL  " & ", Status = 'N' where GUTSCHNR = " & cGutschnr
        
        schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError
    End If
    Label3(1).Caption = Format$(dAltWert, "#####0.00")
    
    'Restbetrag ermitteln
    dNochOffen = dZuZahlen - dAltWert
    
    If dNochOffen < 0 Then
    Label2(2).Caption = "zurück:"
    Else
    Label2(2).Caption = "noch offen:"
    End If
    
    Label3(2).Caption = Format$(dNochOffen, "#####0.00")

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AktualisiereZahlungWK20g"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub LoescheAltenGutscheinWK20g(cLBSatz As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cFeld   As String
    Dim cSQL    As String
    
    cFeld = Left(cLBSatz, 7)
    cFeld = Trim$(cFeld)
    
    cSQL = "Delete from GUTSCH where GUTSCHNR = " & cFeld
    gdBase.Execute cSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoescheAltenGutscheinWK20g"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1

End Sub
Private Sub ReInitDialog20WK20g()
    On Error GoTo LOKAL_ERROR
    
    Dim iFehler As String
    
iFehler = 1
    gbNumTaste = True
iFehler = 2
    frmWKL20!List1.Clear
    frmWKL20!List3.Nodes.Clear
    frmWKL20.Label41(1).Caption = 0
iFehler = 3
    frmWKL20!Label2(6).Caption = "0,00"
    
iFehler = 4
    LeereDialogModul20

    frmWKL20!Text1(0).Text = gcBedienerNr
    gcKreditKarte = ""
    gcZahlMittel = ""
iFehler = 18
    If gbBEDLEER = True Then
        frmWKL20!Label1(8).Caption = ""
        frmWKL20!Text1(0).Text = ""
    End If
iFehler = 21

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ReInitDialog20WK20g"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub SendeDaten2DruckerGutschein2WK20g(lGutschnr As Long, dZuZahlen As Double, dGegeben As Double, dWertRestGutschein As Double)
    On Error GoTo LOKAL_ERROR
    
    Dim lartnr          As Long
    Dim lAnzSatz        As Long
    Dim lAktSatz        As Long
    Dim lAnzZeile       As Long
    Dim lHeute          As Long
    Dim lAnzLbSatz      As Long
    Dim lRet            As Long
    Dim lTask           As Long
    Dim lcount          As Long
    
    Dim cLBSatz         As String
    Dim cFeld           As String
    Dim cDaten          As String
    Dim ctmp            As String
    Dim cTmp2           As String
    Dim cMWST           As String
    Dim cText           As String
    Dim aDeviceName     As String
    Dim cEscapeSequenz  As String
    Dim cWertGutsch     As String
    Dim cArtNr          As String
    Dim cAnz            As String
    
    Dim dAktZeit        As Double
    Dim dNeuZeit        As Double
    Dim dMWSt           As Double
    Dim dRestZahlung    As Double
    Dim dSumGutsch      As Double
    Dim dGRabatt        As Double
    Dim dGRabattWert    As Double
    Dim dSumme          As Double
    Dim dWert           As Double
    Dim dEuro           As Double
    Dim dMWStVoll       As Double
    Dim dMWStErm        As Double
    
    Dim iFileNr         As Integer
    Dim iGesAnzahl      As Long
    Dim iLenZeile       As Integer
    Dim iLevel          As Integer
    Dim iAktCopy        As Integer
    ReDim cDruckZeile(1 To 1) As String
     
    lHeute = Fix(Now)
    If lHeute >= glEUROTAG Then
        gcWaehrung = "EUR"
    Else
        gcWaehrung = gcWaehrung
    End If
    
    dRestZahlung = dZuZahlen
    iLevel = 0
    
    setzedrucker gcBonDrucker
    
    'Drucker an, Display aus, Init Drucker
    aDeviceName = Printer.DeviceName
    cEscapeSequenz = gcInit
    OpenDrawer aDeviceName, cEscapeSequenz
    
SCHUBLADE:
    If gbLadeCom Then
        OpenDrawerViaComPortModul20
    Else
        If gbAPI = False Then
            dAktZeit = Time
            lRet = Shell("Command.com /C " & gcPfad & "LADE.EXE", 6)
            dNeuZeit = Time
            Do While dNeuZeit - dAktZeit < (2 / 86400)
                dNeuZeit = Time
            Loop
        Else
            aDeviceName = Printer.DeviceName
            cEscapeSequenz = gcLade
            OpenDrawer aDeviceName, cEscapeSequenz
        End If
    End If
    
StartPunkt:
    lAnzZeile = 0
    ReDim cDruckZeile(1 To 1) As String
    
    iAktCopy = iAktCopy + 1
    
    'zuerst QR-Code auf dem ersten Bon drucken und dann Papier schneiden <<<<< START
     If iAktCopy > 1 Then
     
        If MitQrCode And E_TSE_Aktiv Then
          'beim altDruckModus kann der Drucker beim Kunde kein QR-Code drucken (alte Drucker)
             If altDruckModus Then
             
             Else
               QRcodeDrucken
             End If
        End If
        
  If altDruckModus Then
    'Papier schneiden (alte Funktion)
        If gbBonDruck Then
            'Kassenbon abschneiden
            If gbAPI = True Then
                aDeviceName = Printer.DeviceName
                cEscapeSequenz = gcSchneiden
                OpenDrawer aDeviceName, cEscapeSequenz
            End If
        End If
    
    Else
        'Papier schneiden (neue Funktion)
        CutPapier
        Sleep 2000
    End If
       
     End If
    'zuerst QR-Code auf dem ersten Bon drucken und dann Papier schneiden <<<<< ENDE
    
    
    iLevel = 1
    cDaten = ""
    iLenZeile = 32
    dSumme = 0
    dMWStVoll = 0
    dMWStErm = 0
    'Artikeldaten an Drucker senden
    lAnzSatz = frmWKL20!List1.ListCount
    
    iLevel = 2
    'Drucker ist bereits auf BonDrucker geschaltet
    aDeviceName = gcBonDrucker
    
    iLevel = 3
    dMWStVoll = 0
    dMWStErm = 0
    
    '***********************************************
    'Drucker ein- und Kundendisplay ausschalten
    '***********************************************
    
    cEscapeSequenz = gcInit
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    'Kopfdaten an Drucker senden
    If gcBild <> "" Then
        aDeviceName = Printer.DeviceName
        cEscapeSequenz = gcBild
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    '************************************************************
    '* 1.Kopfzeile
    '************************************************************

    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
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
    
    '************************************************************
    '* 2.Kopfzeile
    '************************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "Carsten Schröder"
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
    
    '************************************************************
    '* 3.Kopfzeile
    '************************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
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
    
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Artikelpositionen drucken
    '***********************************************
    
    iLevel = 4
    dSumme = 0
    iGesAnzahl = 0
    
    
    
    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        '1.Zeile: ArtNr + MWSTKz + ArtBezeich
        cFeld = Mid(cLBSatz, 7, 6)
        '//2002
        lartnr = CLng(cFeld)
        If cFeld <> "000000" Then
            cDaten = cFeld & " "
            cFeld = Mid(cLBSatz, 72, 1)
            cDaten = cDaten & cFeld & "  "
            cMWST = cFeld
            cFeld = Mid(cLBSatz, 14, 35)
            cFeld = Trim$(cFeld)
            If Len(cFeld) > 17 Then
                cFeld = Left(cFeld, 17)
            End If
            
            '//2002
            If gbDivKosmetik = True Then
                If lartnr = 666666 Then
                    cDaten = cDaten & cFeld
                Else
                    cDaten = cDaten & gcDivKosmetik
                End If
            Else
                cDaten = cDaten & cFeld
            End If
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            ctmp = Mid(cLBSatz, 124, 3)
            If Val(ctmp) > 0 And Left(gFirma.FirmaName, 5) <> "Stief" And gbRabatt Then
                'Zeile nur bei Artikel-Ermäßigung drucken
                
                Dim dArtikelrabattinEuro As Double
                dArtikelrabattinEuro = CDbl(Trim(Mid(cLBSatz, 84, 9)))
                Dim dRabattierterGesamtPreisinEuro As Double
                dRabattierterGesamtPreisinEuro = CDbl(Trim(Mid(cLBSatz, 60, 9)))
                Dim dErgebnisinProz As Double
                dErgebnisinProz = dArtikelrabattinEuro * 100 / (dRabattierterGesamtPreisinEuro + dArtikelrabattinEuro)
                ctmp = Format$(dErgebnisinProz, "###,##0.00")
                
                
                cDaten = "Rabatt:    " & ctmp & " %"
                ctmp = Mid(cLBSatz, 84, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                ctmp = Space(9 - Len(ctmp)) & ctmp
                cDaten = cDaten & ctmp
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            ctmp = Mid(cLBSatz, 1, 5)
            ctmp = Trim$(ctmp)
            ctmp = ctmp & Space$(6 - Len(ctmp))
            cDaten = ctmp & " x"
            
            
            
            
            
            If Left(cLBSatz, 1) = "x" Then
                cAnz = Mid(cLBSatz, 2, 4)
            Else
                cAnz = Mid(cLBSatz, 1, 5)
            End If
            
            
            iGesAnzahl = iGesAnzahl + Val(cAnz)
            
            If gbRabatt Then
                ctmp = Mid(cLBSatz, 74, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                dWert = Val(ctmp)
            Else
                ctmp = Mid(cLBSatz, 50, 9)
                ctmp = fnMoveComma2Point$(ctmp)
                dWert = Val(ctmp)
            End If
            
            If Left(gFirma.FirmaName, 5) <> "Stief" Then
                ctmp = Format$(dWert, "#####0.00")
            Else
                ctmp = Format$((dWert * 100), "########0")
            End If
            ctmp = Space(11 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            
            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = fnMoveComma2Point$(ctmp)
            dWert = Val(ctmp)
            ctmp = Format$(dWert, "#####0.00")
            dSumme = dSumme + dWert
            ctmp = Space(13 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            If cMWST = "V" Then
                dMWSt = dWert / (100 + gdMWStV)
                dMWSt = dMWSt * gdMWStV
                dMWStVoll = dMWStVoll + dMWSt
            ElseIf cMWST = "E" Then
                dMWSt = dWert / (100 + gdMWStE)
                dMWSt = dMWSt * gdMWStE
                dMWStErm = dMWStErm + dMWSt
            Else
                dMWSt = 0
            End If
        Else
'            'Zeile mit Zwischensumme drucken
'            cDaten = "Zwischensumme:     "
'            ctmp = Mid(cLBSatz, 60, 9)
'            ctmp = fnMoveComma2Point$(ctmp)
'            dWert = Val(ctmp)
'            ctmp = Format$(dWert, "#####0.00")
'            ctmp = Space(13 - Len(ctmp)) & ctmp
'            cDaten = cDaten & ctmp

            'Zeile mit Zwischensumme drucken
            ctmp = Mid(cLBSatz, 13, Len(cLBSatz) - 13)
            ctmp = Left(Trim(ctmp), 32)
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    Next lAktSatz
    
    iLevel = 5
    cDaten = String$(iLenZeile, "-")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
      
    If frmWKL20!Label2(3).Visible And Left(gFirma.FirmaName, 5) <> "Stief" And gbRabatt Then
        'Zeile nur bei Gesamt-Ermäßigung drucken
        
        dWert = fnHoleGesamtRabattModul20#()
        
        
        
        
        
        ctmp = frmWKL20!Label2(3).Caption
        
        
        If Len(ctmp) > 6 Then
            ctmp = Left(ctmp, Len(ctmp) - 5)
        End If
        
        Dim cDruckGesRabattBezeichnung As String
            
        cDruckGesRabattBezeichnung = "GesRabatt: "
        cDaten = cDruckGesRabattBezeichnung & Space(5 - Len(ctmp)) & ctmp & "% " & gcWaehrung
        
        
        
        
        
        
        
'        If Len(ctmp) > 3 Then
'            ctmp = Left(ctmp, Len(ctmp) - 3)
'        End If
'
'        If Len(ctmp) > 2 Then
'            cDaten = "Gesamtrabatt:" & Space(3 - Len(ctmp)) & ctmp & "% " & gcWaehrung
'        Else
'            cDaten = "Gesamtrabatt: " & Space(2 - Len(ctmp)) & ctmp & "% " & gcWaehrung
'        End If

        
        ctmp = Format$(dWert, "#####0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***********************************************
        'Zeile Leerzeile drucken
        '***********************************************
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    
    
    'Endbetrag anfang
    
    ctmp = "Endbetrag"
    
    ctmp = Trim$(ctmp)
    ctmp = ctmp & Space$(17 - Len(ctmp))
    iLevel = 605
    
    ctmp = ctmp & Space$(1) & gcWaehrung
    cDaten = ctmp
    gdSumme = dZuZahlen
    ctmp = Format$(gdSumme, "###,##0.00")
    ctmp = Space$(11 - Len(ctmp)) & ctmp
    iLevel = 6102
    cDaten = cDaten & ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    
    iLevel = 6103
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************
    cDaten = String$(iLenZeile, "_")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    cDaten = String$(iLenZeile, "_")
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    '***********************************************
    'Zeile Leerzeile drucken
    '***********************************************
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    'Endbetrag Ende
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    iLevel = 6
'    '//new
'    If (gcWaehrung = "DEM" And dZuZahlen >= 200) Or (gcWaehrung = "EUR" And dZuZahlen >= 100) Then
'        ctmp = "Summe ohne MwSt.:" & Space$(1) & gcWaehrung
'        cDaten = ctmp
'        ctmp = Format$((dZuZahlen - (dMWStVoll + dMWStErm)), "######0.00")
'
'        gdSumme = dZuZahlen
'        ctmp = Space$(11 - Len(ctmp)) & ctmp
'        cDaten = cDaten & ctmp
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'    Else
'        ctmp = "Summe incl. Mwst.:" & Space$(0) & gcWaehrung  '//2002
'        cDaten = ctmp
'        ctmp = Format$(dZuZahlen, "######0.00")
'
'        gdSumme = dZuZahlen
'        ctmp = Space$(11 - Len(ctmp)) & ctmp
'        cDaten = cDaten & ctmp
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'    End If
    
'''    '***********************************************
'''    'MWSt-Beträge andrucken
'''    '***********************************************
'''    ctmp = "MWSt.-Anteil: " & Format$(gdMWStV, "#0") & "%" & Space$(1) & gcWaehrung
'''    cDaten = ctmp
'''    ctmp = Format$(dMWStVoll, "#####0.00")
'''    ctmp = Space$(11 - Len(ctmp)) & ctmp
'''    cDaten = cDaten & ctmp
'''    KonvertAnsiAscii cDaten
'''    cEscapeSequenz = cDaten & vbCrLf
'''
'''    lAnzZeile = lAnzZeile + 1
'''    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'''    cDruckZeile(lAnzZeile) = cEscapeSequenz
'''
'''    ctmp = "MWSt.-Anteil: " & Format$(gdMWStE, "#0") & "%" & Space$(2) & gcWaehrung
'''    cDaten = ctmp
'''    ctmp = Format$(dMWStErm, "#####0.00")
'''    ctmp = Space$(11 - Len(ctmp)) & ctmp
'''    cDaten = cDaten & ctmp
'''    KonvertAnsiAscii cDaten
'''    cEscapeSequenz = cDaten & vbCrLf
'''
'''    lAnzZeile = lAnzZeile + 1
'''    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'''    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
''    '//SUMME Brutto
''    If (gcWaehrung = "DEM" And dZuZahlen >= 200) Or (gcWaehrung = "EUR" And dZuZahlen >= 100) Then
''        ctmp = "Summe incl. MwSt.:" & Space$(0) & gcWaehrung
''        cDaten = ctmp
''        ctmp = Format$(dZuZahlen, "######0.00")
''
''        gdSumme = dZuZahlen
''
''        ctmp = Space$(11 - Len(ctmp)) & ctmp
''        cDaten = cDaten & ctmp
''        KonvertAnsiAscii cDaten
''        cEscapeSequenz = cDaten & vbCrLf
''        lAnzZeile = lAnzZeile + 1
''        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
''        cDruckZeile(lAnzZeile) = cEscapeSequenz
''    End If
''    '//End SUMME Brutto
    
    '***********************************************
    'alle eingereichten Gutscheine auf Kassenbon drucken
    '***********************************************

    dSumGutsch = 0
    lAnzSatz = List3.ListCount

    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = List3.list(lAktSatz)
        cFeld = Left(cLBSatz, 8)
        cFeld = Trim$(cFeld)
        cFeld = Space$(8 - Len(cFeld)) & cFeld


        cFeld = Mid(cLBSatz, 21, 10)
        cFeld = Trim$(cFeld)
        cWertGutsch = cFeld
        cWertGutsch = fnMoveComma2Point$(cWertGutsch)

'        If iAktCopy = 1 Then
            dSumGutsch = dSumGutsch + Val(cWertGutsch)
'        End If

        
        
        
    Next lAktSatz
    
    If dSumGutsch > 0 Then 'Gegeben Gutschein
        ctmp = "Geg Gutscheine:" & Space$(3) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dSumGutsch, "######0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If

    
    '***********************************************
    'Restsumme andrucken
    '***********************************************
    dRestZahlung = dZuZahlen - dSumGutsch
    If dRestZahlung > 0 Then
        If dZuZahlen = 0 Then
            dRestZahlung = 0
            gdGutLastRest = dRestZahlung
        Else
            ctmp = "Restzahlung:" & Space$(6) & gcWaehrung
            cDaten = ctmp
            If dSumGutsch = 0 Then
                ctmp = Format$(dSumGutsch, "######0.00")
            Else
                If gbGutschUNDlastschrift = True Then
                    gdGutLastRest = dRestZahlung
                    ctmp = Format$(gdGutLastRest, "######0.00")
                Else
                    ctmp = Format$(dRestZahlung, "######0.00")
                End If
            End If
            ctmp = Trim$(ctmp)
            ctmp = Space$(11 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    Else
        dSumGutsch = 0
        dRestZahlung = dZuZahlen - (dSumGutsch)
        ctmp = "Restzahlung:" & Space$(6) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dSumGutsch, "######0.00")
        ctmp = Trim$(ctmp)
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'zusätzlich gegebenen Betrag andrucken
    '***********************************************
    If dGegeben > 0 Then
        ctmp = "Gegeben ("
        
        If UCase(gcZahlMittel) = "KA" Then
            ctmp = ctmp & gcKreditKarte
        Else
            ctmp = ctmp & gcZahlMittel
        End If
        
        ctmp = ctmp & "): " & Space$(4) & gcWaehrung
        cDaten = ctmp
        ctmp = Format$(dGegeben, "######0.00")
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Rückgeld andrucken
    '***********************************************
    If gcRueckgeld <> "0,00" Then
        ctmp = "Zurück(BA):" & Space$(7) & gcWaehrung
        cDaten = ctmp
        ctmp = gcRueckgeld
        ctmp = Trim$(ctmp)
        ctmp = Space$(11 - Len(ctmp)) & ctmp
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        
        
        
    End If
    
    '***********************************************
    'Rückgutschein andrucken
    '***********************************************
    If gcRueckGutsch <> "" Then
        cDaten = gcRueckGutsch
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    
    
    
    
    'neuer Anfang
    
    '***********************************************
    'Zeile Leerzeile drucken
    '***********************************************
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    '***********************************************
    'Zeile 'Anzahl Artikel' drucken
    '***********************************************
    If iGesAnzahl > 1 Then
        ctmp = "Anzahl Artikel: " & iGesAnzahl
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    '***********************************************
    'Bedienernamen andrucken
    '***********************************************
    ctmp = "Es bediente Sie"
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    ctmp = gcBediener
    cDaten = ctmp
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    If gbKASSNRUNTER = False Then
        '***********************************************
        'Zeile 'Kassennummer' drucken
        '***********************************************
        
        ctmp = "Kasse: " & gcKasNum
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    End If
    
    
    
    
    '***********************************************
    'Kundenname + Kundennr andrucken
    '***********************************************
    If frmWKL20!Label2(7).Caption <> "0" Then
    
    
        ctmp = "Ihre KundenNr: " & frmWKL20!Label2(7).Caption
        cDaten = ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
'        ctmp = frmWKL20!Label2(7).Caption
        
        
        If gbKUNDENA = True Then
        
            If gbKUIBONfirma Then
                ctmp = lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).firma
            
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
                ctmp = lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).titel
            
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
                ctmp = lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).vorname
            End If
            
            If gbKUIBONname Then
                If ctmp = "" Then
                    ctmp = ctmp & ""
                Else
                    ctmp = ctmp & " "
                End If
                ctmp = ctmp & lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).nachname
            End If
    
        
            iLevel = 614
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
                ctmp = lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).strasse
            
                iLevel = 615
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
                ctmp = lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).Plz
            End If
            
            If gbKUIBONort Then
                If ctmp = "" Then
                    ctmp = ctmp & ""
                Else
                    ctmp = ctmp & " "
                End If
                ctmp = ctmp & lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).Ort
            End If
        
            iLevel = 616
            If Len(ctmp) > 32 Then
                ctmp = Left(ctmp, 32)
            End If
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            iLevel = 617
            
            If gbKUIBONtel Then
                ctmp = lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).telefon
                If ctmp <> "" Then
                    cDaten = "Tel " & ctmp
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
            End If
            
            iLevel = 618
            
            If gbKUIBONmobil Then
                ctmp = lookingForKundendaten(Trim(frmWKL20!Label2(7).Caption)).Mobiltel
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
        
        
        
        
        
'        ctmp = fnHoleKundenNameMOD1(ctmp)
'
'        If Len(ctmp) > 32 Then
'            ctmp = Left(ctmp, 32)
'        End If
'        cDaten = ctmp
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        'wenn schon mit Kundenbindung dann auch mit bonusvariante 1
        
        Dim sEinText As String
        Dim sWort As String
        
        cLBSatz = ""
        
        If giBonusNr = 0 Then
        
            sEinText = Trim(gsTextVor) & " " & Val(frmWKL20!Label1(14).Caption) & " " & Trim(gsTextNach) & " "
'            sEinText = Trim(gsTextVor) & " " & Val(lookingForKundendaten(Trim(Label2(7).Caption)).BONUS) & " " & Trim(gsTextNach) & " "
        
            If Len(sEinText) > iLenZeile Then
                Do While Len(sEinText) > 0
                    sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
                    
                    If Len(cLBSatz & sWort & Space(1)) > iLenZeile Then
                    
                        cDaten = cLBSatz
                        KonvertAnsiAscii cDaten
                        cEscapeSequenz = cDaten & vbCrLf
                        
                        lAnzZeile = lAnzZeile + 1
                        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                        cDruckZeile(lAnzZeile) = cEscapeSequenz
                        
                        cLBSatz = ""
                        
                    End If
                    
                    cLBSatz = cLBSatz & sWort & Space(1)
                    sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
                Loop
                
                cDaten = cLBSatz
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                If Val(gsWWBonusGDAUER) > 0 Then
                    cDaten = "Gültigkeit bis " & Format(DateValue(Now) + gsWWBonusGDAUER, "DD.MM.YYYY")
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
                
            Else
                cLBSatz = sEinText
                cDaten = cLBSatz
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
        End If
        
    End If
    
    '***********************************************
    'Zeile Datum, BelegNr, Uhrzeit drucken
    '***********************************************
    iLevel = 615
    ctmp = Format$(Date, "DD.MM.YYYY")
    cDaten = ctmp
    
    iLevel = 6154
    ctmp = Format$(Now, "HH:MM")
    iLevel = 6155
    cDaten = cDaten & Space$(4) & ctmp
    
    If giZahlArt = giKOLLEGE Then
        ctmp = "0"
    Else
        ctmp = Format$(gdBonNr, "#####0")
    End If
    iLevel = 6151
    ctmp = gcKasNum & "/" & ctmp
    iLevel = 6152
    ctmp = Space$(8 - Len(ctmp)) & ctmp
    iLevel = 6153
    cDaten = cDaten & Space$(4) & ctmp
    
    iLevel = 6156
    KonvertAnsiAscii cDaten
    iLevel = 6157
    cEscapeSequenz = cDaten & vbCrLf
    iLevel = 6158
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    iLevel = 6159
    cDruckZeile(lAnzZeile) = cEscapeSequenz

    iLevel = 7
    
    '***********************************************
    'Zeile Leerzeile drucken
    '***********************************************
    cEscapeSequenz = vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    If gcRueckgeld = "" Then gcRueckgeld = "0,00"
    
    
    'bei eventueller Gutschein_Barauszahlung ein Unterschriftenfeld erzeugen
    If gbGUTSCHBARAUSZAHLUNGMITUNTER = True And gcRueckgeld <> "0,00" Then
    
        cDaten = "Gutscheinauszahlung erhalten:"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***********************************************
        'Zeile Leerzeile drucken
        '***********************************************
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '***********************************************
        'Zeile Leerzeile drucken
        '***********************************************
        cEscapeSequenz = vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        '***********************************************
        'Zeile Trennstrich drucken
        '***********************************************
        cDaten = String$(iLenZeile, "_")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        cDaten = "(Unterschrift)"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    '****enventuell angefallene Garantiedaten einfügen 'Achtung 3 Mal im Programm
    
    If sind_Garatie_daten_zu_drucken(gdBonNr) = True Then
    
        
        cDaten = "*** Garantie - Informationen ***"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        ermGarantie_daten gdBonNr
        
        '3 Arrays auslesen
        Dim cWert As String
        Dim iCount As Integer
        For iCount = 1 To UBound(gcArrArtNr)
        
        
            'Artikelnummer
            cWert = gcArrArtNr(iCount)
            If cWert <> "" Then
                
            
                cDaten = "zu Artikel: " & cWert
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            End If
            
            'Seriennummer
            
            cWert = gcArrSerienNr(iCount)
            If cWert <> "" Then
                
                If Len(cWert) <= 22 Then
            
                    cDaten = "SerienNr: " & cWert
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                ElseIf Len(cWert) <= 32 Then
                
                    cDaten = "SerienNr: "
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    cDaten = cWert
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                ElseIf Len(cWert) > 32 Then
                
                    Dim sNeuWert As String
                    
                    sNeuWert = Right(cWert, Len(cWert) - 22)
                
                    cDaten = "SerienNr: " & Left(cWert, 22)
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    cDaten = Left(sNeuWert, 32)
                    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    Do While Len(sNeuWert) >= 32
                        sNeuWert = Right(sNeuWert, Len(sNeuWert) - 32)
                        
                        If sNeuWert <> "" Then
                            cDaten = Left(sNeuWert, 32)
                            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                            KonvertAnsiAscii cDaten
                            cEscapeSequenz = cDaten & vbCrLf
                            lAnzZeile = lAnzZeile + 1
                            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                            cDruckZeile(lAnzZeile) = cEscapeSequenz
                        End If
                    Loop
                    
                    
                
                End If
            
            End If
            
            
        Next iCount
        
        
    
        cDaten = "*** Garantie - Informationen ***"
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    End If
    
    
    
    '**** ENDE enventuell angefallene Garantiedate einfügen
    
    
    
    
    
    
    
    '***********************************************
    'Trennzeile drucken
    '***********************************************
    cDaten = String$(iLenZeile, gsSTERNZEICH)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    If giPreisKz = 6 Then
        '***********************************************
        'Zeile Nettoumsatz
        '***********************************************
        iLevel = 6
        ctmp = "Nettoumsatz"
        cDaten = ctmp
        ctmp = Format$(gdSumme, "###,##0.00")
        ctmp = Space$(15 - Len(ctmp)) & ctmp
        iLevel = 601
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    Else
    
        '***********************************************
        'Zeile Nettoumsatz
        '***********************************************
        iLevel = 6
        ctmp = "Nettoumsatz"
        cDaten = ctmp
        ctmp = Format$((gdSumme - (dMWStVoll + dMWStErm)), "###,##0.00")
        ctmp = Space$(15 - Len(ctmp)) & ctmp
        iLevel = 601
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        '***********************************************
        'Zeile volle MWSt drucken
        '***********************************************
        If dMWStVoll <> 0 Then
            iLevel = 609
        
            ctmp = "MWSt.-Anteil: " & Format$(gdMWStV, "#0") & "%"
            cDaten = ctmp
            ctmp = Format$(dMWStVoll, "###,##0.00")
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    
        '***********************************************
        'Zeile erm. MWSt drucken
        '***********************************************
        If dMWStErm <> 0 Then
            iLevel = 610
        
            ctmp = "MWSt.-Anteil: " & Format$(gdMWStE, "#0") & "%"
            cDaten = ctmp
            ctmp = Format$(dMWStErm, "###,##0.00")
            ctmp = Space$(10 - Len(ctmp)) & ctmp
            cDaten = cDaten & ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    
        '***********************************************
        'Zeile Bruttoumsatz
        '***********************************************
        
        iLevel = 6101
        ctmp = "Bruttoumsatz"
        cDaten = ctmp
        ctmp = Format$(gdSumme, "###,##0.00")
        ctmp = Space$(14 - Len(ctmp)) & ctmp
        iLevel = 6102
        cDaten = cDaten & ctmp
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        
        iLevel = 6103
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    End If
    
    
    
    
    
    
    
    
    
    
    

    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************

    cDaten = String$(iLenZeile, gsSTERNZEICH)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    
    '***********************************************
    'alle eingereichten Gutscheine auf Kassenbon drucken
    '***********************************************







    lAnzSatz = List3.ListCount

    For lAktSatz = 0 To lAnzSatz - 1
        cLBSatz = List3.list(lAktSatz)
        cFeld = Left(cLBSatz, 8)
        cFeld = Trim$(cFeld)
        cFeld = Space$(8 - Len(cFeld)) & cFeld

        If gbmGDetails Then
            Dim cVKdAT As String
            Dim cKK_art As String

            cVKdAT = ermGutschVkdat(cFeld)

            If cVKdAT <> "" Then
                cVKdAT = Format$(cVKdAT, "DD.MM.YYYY")
            Else
                cVKdAT = ""
            End If

            cKK_art = ermGutschkkart(cFeld, cVKdAT)
        End If



        cDaten = "Gutschein" & cFeld & " " & gcWaehrung & " "
        cFeld = Mid(cLBSatz, 21, 10)
        cFeld = Trim$(cFeld)
'''        cWertGutsch = cFeld
'''        cWertGutsch = fnMoveComma2Point$(cWertGutsch)
'''
'''        If iAktCopy = 1 Then
'''            dSumGutsch = dSumGutsch + Val(cWertGutsch)
'''        End If
        cFeld = Space$(10 - Len(cFeld)) & cFeld
        cDaten = cDaten & cFeld
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
        
        
        
        If gbmGDetails Then
            ctmp = cVKdAT & Space(2) & cKK_art
            ctmp = Space$(7) & ctmp
            cDaten = ctmp
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
        
        
        
    Next lAktSatz
    
    
    
    '***********************************************
    'Zeile Trennstrich drucken
    '***********************************************

    cDaten = String$(iLenZeile, gsSTERNZEICH)
    cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
    KonvertAnsiAscii cDaten
    cEscapeSequenz = cDaten & vbCrLf
    lAnzZeile = lAnzZeile + 1
    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    
    
    'adt beginn
    If gbADTBON Then
        With AktEcashBon
            
            If .Kartenart <> "" Then
                cDaten = .Funktion & " " & .Kartenart
            Else
                cDaten = .Funktion
            End If
            
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
'            cDaten = .Kartenart
'            KonvertAnsiAscii cDaten
'            cEscapeSequenz = cDaten & vbCrLf
'            lAnzZeile = lAnzZeile + 1
'            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            If .Storno <> "" Then
                cDaten = .Storno
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .Betrag <> "" Then
                cDaten = .Betrag
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            cDaten = .Datum & "   " & .Uhrzeit
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            

            If .TerminalID <> "" Then
                cDaten = .TerminalID
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            
            If .Tracenummer <> "" Then
                cDaten = "TA Nr.:       " & .Tracenummer
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .TracenummerSTORNO <> "" Then
                cDaten = "TA Nr.(alt):  " & .TracenummerSTORNO
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .Belegnummer <> "" Then
                cDaten = "Belegnr.:     " & .Belegnummer
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .BLZ <> "" Then
                cDaten = "Bankleitzahl: " & .BLZ
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .BLZ <> "" Then
                cDaten = "Kontonr.:     " & .Kontonummer
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .Karte <> "" Then
                cDaten = "Karte:        " & .Karte
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            cDaten = .Kartenfolgenummer & "  " & .Verfallsdatum & "  " & .AIDParameter
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
'            cDaten = .Verfallsdatum
'            KonvertAnsiAscii cDaten
'            cEscapeSequenz = cDaten & vbCrLf
'            lAnzZeile = lAnzZeile + 1
'            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
'            cDaten = .AIDParameter
'            KonvertAnsiAscii cDaten
'            cEscapeSequenz = cDaten & vbCrLf
'            lAnzZeile = lAnzZeile + 1
'            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'            cDruckZeile(lAnzZeile) = cEscapeSequenz

            If .Autorisierungsmerkmal <> "" Then
                cDaten = .Autorisierungsmerkmal
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .Referenzparameter <> "" Then
                cDaten = .Referenzparameter
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .ReferenzNr <> "" Then
                cDaten = .ReferenzNr
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .VuNr <> "" Then
                cDaten = .VuNr
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            
            cDaten = .Online & " " & .Manuell & " " & .TelefonBuchung
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            
            If .ProviderText_01 <> "" Then
                cDaten = .ProviderText_01
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .ProviderText_02 <> "" Then
                cDaten = .ProviderText_02
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .ProviderText_03 <> "" Then
                cDaten = .ProviderText_03
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            If .ProviderText_04 <> "" Then
                cDaten = .ProviderText_04
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            
            
            
'            '****Hier Frage ob ELV/POZ - Text nötig

            If bErmachtigung = True Then
            
                cDaten = ""
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            
                cDaten = "Ermächtigung Lastschrifteinzug:"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = ""
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "Ich ermächtige hiermit das"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "oben genannte Unternehmen,"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "den als Endsumme ausgewiesenen"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "Betrag von meinem durch Bank- "
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "leitzahl und Kontonummer be-"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "zeichneten Konto durch Last-"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "schrift einzuziehen."
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = ""
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz

                cDaten = "Ermächtigung Adressweitergabe:"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = ""
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "Ich weise mein Kreditinstitut,"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "das durch die Bankleitzahl "
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "bezeichnet ist, unwiderruflich"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "an, bei Nichteinlösung der Last-"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "schrift oder bei Widerspruch"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "gegen die Lastschrift dem"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "Unternehmen oder einem von"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "ihm beauftragten Dritten auf "
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "Anforderung meinen Namen und"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
    
                cDaten = "meine Adresse mitzuteilen, damit"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "das Unternehmen seinen Anspruch"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = "gegen mich geltend machen kann."
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
            End If
            '****Hier Frage ob ELV/POZ - Text nötig ENDE
            '****Hier Frage ob Unterschrift nötig
            If bUnterschrift Then
            
                cDaten = "Unterschrift:"
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = ""
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = ""
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                cDaten = ""
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            '****Hier Frage ob Unterschrift nötig ENDE
            
            
            If Left(.ErgebnisText_1, 2) = "**" Then
                cDaten = .ErgebnisText_1
                cDaten = .ErgebnisText_1
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            End If
            
            cDaten = .ErgebnisText_2
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
            
            
            '***********************************************
            'Zeile Trennstrich drucken
            '***********************************************
            cDaten = String$(iLenZeile, gsSTERNZEICH)
'            cDaten = String$(iLenZeile, "*")
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
            
        End With
        If iAktCopy < 2 Then
            gbADTBON = True
        Else
            gbADTBON = False
        End If
    End If
    'adt ende
    
    
    
    
    
    
    
    
    
  'TSE Footer START  <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< START
    
    'wenn Endbetrag <= 0, dann TSE überspringen, weil ein
    'Kassenbon mit 0 Endbetrag nicht signiert wird
    If gdSumme > 0 Then
    
    Else
    
        R_StartTime = ""
        R_FinishTime = ""
        R_TransactionNr = ""
        R_QRCodeAlsText = ""
        R_QRCodeAlsImgPath = ""
        R_FinishSignatur = ""
        R_StartSignatur = ""
        GoTo NACH_TSE
        
    End If
    

    If E_TSE_Aktiv And TSE_OK Then
 
            'nur erster Kassenbon wird signiert
            If iAktCopy < 2 Then
                  
             'Bon signieren <<<<<<<<<<<< START
                      
               TransactionSchreiben "", 1, 1, dMWStVoll, dMWStErm, 0, 0, 0, 0, gdSumme
            
             'Bon signieren <<<<<<<<<<<< ENDE
             
            End If
      
            If TSE_OK Then

                    '''''''''''''''''''  TSE Start  ''''''''''''''''''''''''
                    cDaten = "TSE Start: " & R_StartTime

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                    '''''''''''''''''''  TSE Ende  ''''''''''''''''''''''''''
                    cDaten = "TSE Ende: " & R_FinishTime

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                    '''''''''''''''  TSE Transaction.Nr  '''''''''''''''''''
                    cDaten = "TSE Transaction.Nr: " & R_TransactionNr

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz

                    '''''''''''''''''''  TSE Signatur  ''''''''''''''''''''''''
                    cDaten = "TSE Signatur: " & vbNewLine & SplitStringNachCharZahl(5, R_FinishSignatur, 32)

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    

                    '''''''''''''''''''  TSE alle Info zusammen (Optional) ''''''''''''''''''''''''
'                    cDaten = "TSE Info: " & vbNewLine & SplitStringNachCharZahl(5, R_QRCodeAlsText, 32)
'
'                    KonvertAnsiAscii cDaten
'                    cEscapeSequenz = cDaten & vbCrLf
'
'                    lAnzZeile = lAnzZeile + 1
'                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
               Else

                    cDaten = "TSE nicht erreichbar !!!"

                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf

                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
               End If
        Else

            cDaten = "TSE ist deaktiviert/falsch" & vbNewLine & "     initialisiert !!!"

            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf

            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        
'        'TSS Footer
'
'        cDaten = "TSE Start: " & TSS.TRStart 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'
'
'        cDaten = "TSE Ende: " & TSS.TRFinish 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'
'
'        cDaten = "TSE TaNr: " & TSS.TRNo 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
'
'
'
'
'
'        cDaten = "TSE Serial: " & TSS.Serial 'Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
'
'        KonvertAnsiAscii cDaten
'        cEscapeSequenz = cDaten & vbCrLf
'
'        lAnzZeile = lAnzZeile + 1
'        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'        cDruckZeile(lAnzZeile) = cEscapeSequenz
         
        
    End If
    
    'TSE Footer ENDE <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ENDE
    
    
        '***********************************************
        'Zeile Trennstrich drucken
        '***********************************************
    
        cDaten = String$(iLenZeile, gsSTERNZEICH)
    '    cDaten = String$(iLenZeile, "*")
        cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
        KonvertAnsiAscii cDaten
        cEscapeSequenz = cDaten & vbCrLf
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        cDruckZeile(lAnzZeile) = cEscapeSequenz
    
NACH_TSE:
    
    
    '**********************************************************
    '* 1.Fußzeile
    '**********************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "KEIN GÜLTIGER KASSENBON!"
    Else
        cDaten = gcBonText(2)
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
    
    '**********************************************************
    '* 2.Fußzeile
    '**********************************************************
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(3)
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
    
    '**********************************************************
    '* 3.Fußzeile
    '**********************************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(5)
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
    'Fußzeile 4 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(6)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    '***********************************************
    'Fußzeile 5 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(7)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 6 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(8)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If

    '***********************************************
    'Fußzeile 7 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(9)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 8 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(10)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    '***********************************************
    'Fußzeile 9 drucken
    '***********************************************
    
    If gbDEMO Then
        'HIER FÜR DEMO FESTTEXT
        cDaten = "DEMO - VERSION!"
    Else
        cDaten = gcBonText(11)
    End If
    
    If Trim$(cDaten) <> "" Then
        cDaten = Trim$(cDaten)
        If cDaten <> "" Then
            cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
            KonvertAnsiAscii cDaten
            cEscapeSequenz = cDaten & vbCrLf
            lAnzZeile = lAnzZeile + 1
            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
            cDruckZeile(lAnzZeile) = cEscapeSequenz
        End If
    End If
    
    iLevel = 10
    
    
    
    'Am Ende eventuell einen Rabattgutschein für den nächsten Einkauf
    
    iLevel = 1050
    
    cLBSatz = ""
        
    Dim bGutscheinbedingung_erfuellt As Boolean
    bGutscheinbedingung_erfuellt = False
            
    If giBonusNr = 1 Then
        iLevel = 1052
        If gbWWKundBi Then 'nur mit Kundenbindung Gutschein anbieten
            If frmWKL20!Label2(7).Caption <> "0" And frmWKL20!Label2(7).Visible Then
                bGutscheinbedingung_erfuellt = True
            End If
        Else
            bGutscheinbedingung_erfuellt = True
        End If
        
        iLevel = 1053
    
        If bGutscheinbedingung_erfuellt = True Then
        
            
            iLevel = 1054
            'für alle bonusfähigen Artikel größer Schwellenwert gibt es einen Rabattgutschein
            Dim dRabattfWert As Double
            Dim dDruckRabattWert As Double
            
            dRabattfWert = ermWertrabattf_Artikel
            
            'ich drucke Gutschein - dafür nehme ich den Bonus beim Kunden zurück
                           
            iLevel = 1055
            BonusVeränderung "negativ", CLng(frmWKL20!Label2(7).Caption), dRabattfWert, 0
            iLevel = 1056
            If dRabattfWert >= CDbl(gsWWSchwellenwert) Then
                '***********************************************
                'Zeile Leerzeile drucken
                '***********************************************
                cDaten = String$(iLenZeile, " ")
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                '***********************************************
                'Zeile  drucken
                '***********************************************
                
                cDaten = "---------- Gutschein -----------"
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
                'Zeile Leerzeile drucken
                '***********************************************
                cDaten = String$(iLenZeile, " ")
                cDaten = Space$((iLenZeile - Len(cDaten)) / 2) & cDaten
                KonvertAnsiAscii cDaten
                cEscapeSequenz = cDaten & vbCrLf
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
                
                iLevel = 1057
            
                If gsWWArt = "Prozent" Then
                    iLevel = 1058
                    dDruckRabattWert = dRabattfWert * CDbl(gsWWwert) / 100
                Else
                    iLevel = 1059
                    dDruckRabattWert = CDbl(gsWWwert)
                End If
                iLevel = 1060
                sEinText = Trim(gsTextVor) & " " & Format(dDruckRabattWert, "###,##0.00") & " " & Trim(gsWWZeichen) & " " & Trim(gsTextNach) & " "
            
                If Len(sEinText) > iLenZeile Then
                    Do While Len(sEinText) > 0
                        sWort = Mid(sEinText, 1, InStr(1, sEinText, " ") - 1)
                        
                        If Len(cLBSatz & sWort & Space(1)) > iLenZeile Then
                            cDaten = cLBSatz
                            KonvertAnsiAscii cDaten
                            cEscapeSequenz = cDaten & vbCrLf
                            
                            lAnzZeile = lAnzZeile + 1
                            ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                            cDruckZeile(lAnzZeile) = cEscapeSequenz
                            cLBSatz = ""
                        End If
                        
                        iLevel = 1061
                        
                        cLBSatz = cLBSatz & sWort & Space(1)
                        sEinText = Mid(sEinText, Len(sWort) + 2, Len(sEinText) - Len(sWort) + 1)
                    Loop
                    
                    cDaten = cLBSatz
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                    
                    If Val(gsWWBonusGDAUER) > 0 Then
                        cDaten = "Gültigkeit bis " & Format(DateValue(Now) + gsWWBonusGDAUER, "DD.MM.YYYY")
                        KonvertAnsiAscii cDaten
                        cEscapeSequenz = cDaten & vbCrLf
                        lAnzZeile = lAnzZeile + 1
                        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                        cDruckZeile(lAnzZeile) = cEscapeSequenz
                    End If
                    
                Else
                
                    iLevel = 1062
                    cLBSatz = sEinText
                    cDaten = cLBSatz
                    KonvertAnsiAscii cDaten
                    cEscapeSequenz = cDaten & vbCrLf
                    lAnzZeile = lAnzZeile + 1
                    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                    cDruckZeile(lAnzZeile) = cEscapeSequenz
                End If
                


            End If
        End If
    End If
    
    
    
    
    
    
    
    
    '***********************************************
    'ein paar Leerzeilen drucken  <<<<<<<<<<<< START
    '***********************************************
    If Not MitQrCode Or Not E_TSE_Aktiv Or Not gdSumme > 0 Or altDruckModus Then
    
        'Barcode Bonus auf Bon
        If gsWWBonusArtnr <> "0" And dDruckRabattWert > 0 Then
    
        Else
    
    
            For lcount = 1 To gbLeereZeil
                If lcount = gbLeereZeil Then
                    cEscapeSequenz = vbCrLf
                Else
                    cEscapeSequenz = " " & vbCrLf
                End If
                lAnzZeile = lAnzZeile + 1
                ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
                cDruckZeile(lAnzZeile) = cEscapeSequenz
            Next lcount
    
        End If

    End If
    '***********************************************
    'ein paar Leerzeilen drucken  <<<<<<<<<<<< ENDE
    '***********************************************
    
    'Schublade nur einmal öffnen
    
    If iAktCopy = 2 Then
        GoTo BON_DRUCKEN
    End If
    iLevel = 12
    
    '//Trotz BON-NEIN wird es in KASSBON für 2.Bon gespeichert
    If gbBonDruck Then
BON_DRUCKEN:
        If gbAPI = True Then
            OpenDrawer3 aDeviceName, cDruckZeile(), lAnzZeile
        Else
            OpenDrawer4 aDeviceName, cDruckZeile(), lAnzZeile
        End If
        
        
        
        Dim bPlusLZ As Boolean
        bPlusLZ = False
        
        If gbBonDruck Then
            If gsWWBonusArtnr <> "0" Then
                If dDruckRabattWert > 0 Then
                    Barcode_Bonus CStr(dDruckRabattWert), "7"
                    bPlusLZ = True
                End If
    
            End If
            
        Else
        
            If gsWWBonusArtnr <> "0" And dDruckRabattWert > 0 Then
                    
                bPlusLZ = True
            End If
        
        End If
        
        
        
        
        
        
        If iAktCopy = 1 Then
            'Bon-Daten sichern
            
            
            Dim cKundnr As String
            cKundnr = ""
            If frmWKL20!Label2(7).Caption <> "0" Then
                cKundnr = frmWKL20!Label2(7).Caption
            End If
            
            SichernBonDaten cDruckZeile(), lAnzZeile, "", cKundnr, False
        End If
        
'BON_SCHNEIDEN:
'        If gbAPI Then
'            aDeviceName = Printer.DeviceName
'            cEscapeSequenz = gcSchneiden
'            If gbAPI = True Then
'                OpenDrawer aDeviceName, cEscapeSequenz
'            End If
'
'            iLevel = 11
'        End If
        
ZWEITER_BON:
        If iAktCopy < 2 Then
            If gb2BONGUVK Then
                GoTo StartPunkt
            End If
        End If
        
        Erase cDruckZeile
        
        
   'uncommit die folgende Zeile zum Drucken des QR-Codes auf dem Bon
    If MitQrCode And E_TSE_Aktiv Then
      'beim altDruckModus kann der Drucker beim Kunde kein QR-Code drucken (alte Drucker)
       If altDruckModus Then
             
       Else
        QRcodeDrucken
       End If
    End If

If altDruckModus Then
    'Papier schneiden (alte Funktion)
        If gbBonDruck Then
            'Kassenbon abschneiden
            If gbAPI = True Then
                aDeviceName = Printer.DeviceName
                cEscapeSequenz = gcSchneiden
                OpenDrawer aDeviceName, cEscapeSequenz
            End If
        End If
    
    Else
        'Papier schneiden (neue Funktion)
        CutPapier
    End If
    
     
     
GUTSCHEIN:
        'gekaufte Gutscheine drucken
        lAnzLbSatz = frmWKL20!List1.ListCount
        For lcount = 0 To lAnzLbSatz - 1
            cLBSatz = frmWKL20!List1.list(lcount)
            cArtNr = Mid(cLBSatz, 7, 6)
            If cArtNr = "666666" Then
            
                 If Not altDruckModus Then
                    PaarLeereZeilenDrucken
                    Sleep 2000
                 End If
                
                DruckeGutscheinBonModul20 cLBSatz
                
                If altDruckModus Then
                    'Papier schneiden (alte Funktion)
                        If gbBonDruck Then
                            'Kassenbon abschneiden
                            If gbAPI = True Then
                                aDeviceName = Printer.DeviceName
                                cEscapeSequenz = gcSchneiden
                                OpenDrawer aDeviceName, cEscapeSequenz
                            End If
                        End If
                    
                 Else
                        'Papier schneiden (neue Funktion)
                        CutPapier
                 End If
                
            End If
        Next lcount
        
        If lGutschnr > 0 Then
            cLBSatz = Space$(100)
            'Rückgabe-Gutschein drucken
            ctmp = Trim$(Str$(lGutschnr))
            ctmp = Space$(8 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 24, 8) = ctmp
            
            ctmp = Format$(dWertRestGutschein, "#####0.00")
            ctmp = Trim$(ctmp)
            ctmp = Space$(9 - Len(ctmp)) & ctmp
            Mid(cLBSatz, 60, 9) = ctmp
            DruckeGutscheinBonModul20 cLBSatz
        End If
    Else '//Trotz BON-NEIN...
        frmWKL20!Label8(0).Caption = ""
        frmWKL20!Label8(1).Caption = ""
        SichernBonDaten cDruckZeile(), lAnzZeile, "", "", False
    End If
 
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SendeDaten2DruckerGutschein2WK20g"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
'    Resume Next

End Sub

Private Sub ProtokolliereALTGutscheinWK20g(lkunde As Long, dWert As Double, lGutschnr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lPos        As Long
    Dim cSatz       As String
    Dim cDatum      As String
    Dim czeit       As String
    Dim cKunde      As String
    Dim cBed        As String
    Dim cFil        As String
    Dim cWert       As String
    Dim cBonNr      As String
    Dim cKasnum     As String
    Dim cGutschnr   As String
    Dim cPfad       As String
    Dim iFileNr     As Integer
    Dim cZeile2     As String
    Dim cSatz1      As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"
    
    cDatum = Format$(Now, "DD.MM.YYYY")                 'Datum der Gutschein-Generierung
    czeit = Format$(Now, "HH:MM:SS")                    'Uhrzeit der Gutschein-Generierung
    
    cKasnum = gcKasNum                                  'Kassennummer
    cKasnum = Space$(2 - Len(cKasnum)) & cKasnum
    
    cBed = gcBedienerNr                                 'Bediener
    cBed = Space$(3 - Len(cBed)) & cBed
    
    cFil = gcFilNr                                      'filiale
    cFil = Space$(2 - Len(cFil)) & cFil
    
    cBonNr = Format$(gdBonNr, "#####0")                 'Bon-Nummer
    cBonNr = Space$(6 - Len(cBonNr)) & cBonNr
    
    cKunde = lkunde                                     'Kunde
    cKunde = Space$(10 - Len(cKunde)) & cKunde
    
    cWert = Format$(dWert, "######0.00")                'Wert
    cWert = Space$(10 - Len(cWert)) & cWert
    
    cGutschnr = Format$(lGutschnr, "#######0")          'Gutscheinnummer
    cGutschnr = Space$(10 - Len(cGutschnr)) & cGutschnr
    
    
    cSatz1 = "Datum      Uhrzeit   Kasse    Filiale  Bon   Bed "
    cSatz1 = cSatz1 & "       Kunde   Wert      GutschNr "
    cSatz1 = cSatz1 & Chr$(13) & Chr$(10)
    
    cSatz = cDatum & " " & czeit & " " & cKasnum & "       " & cFil & "      " & cBonNr & " " & cBed & " "
    cSatz = cSatz & cKunde & "  " & cWert & " " & cGutschnr
    cSatz = cSatz & Chr$(13) & Chr$(10)
    cSatz = cSatz & Chr$(13) & Chr$(10)
    
    
    
    
    iFileNr = FreeFile
    Open cPfad & "ALT_GUT.TXT" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cPfad & "ALT_GUT.TXT"
        
    End If
    
    Kill cPfad & "ALT_GUT.TXT"
    
    iFileNr = FreeFile
    Open cPfad & "ALT_GUT.TXT" For Binary As #iFileNr
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz1
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cSatz
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "ProtokolliereALTGutscheinWK20g"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
        
        Fehlermeldung1
    End If

End Sub
Private Sub SchreibeAltenGutscheinWK20g()
    On Error GoTo LOKAL_ERROR
    
    Dim lGutschNrNeu  As Long
    Dim lGutschnr     As Long
    Dim lGutschNrMax  As Long
    Dim dWert         As Double
    Dim cFeld         As String
    Dim cFiliale      As String
    Dim ctmp          As String
    Dim i             As Integer
    
    '** Daten aus Dialog frmWKL20 **
    Dim lbednu        As Long
    Dim lDatum        As Long
    Dim lKUNDNR       As Long
    Dim lSumme        As Long
    Dim cSQL          As String
    Dim cSQL1         As String
    Dim cLBSatz       As String
    Dim rsrs          As Recordset
    Dim cWert         As String
    
    cWert = "A"
    cFeld = Trim$(Text2.Text)
    If cFeld = "" Then
        MsgBox "Bitte einen Wert eingeben!", vbCritical, "STOP!"
        Text2.SetFocus
        Exit Sub
    End If
    
    cFeld = fnMoveComma2Point$(cFeld)
    dWert = Val(cFeld)
    If InStr(cFeld, ".") = 0 Then
        dWert = dWert / 100
        cFeld = Format$(dWert, "#####0.00")
        Text2.Text = cFeld
    End If
    
    If gbGutsch Then
        frmWKLak.Show 1
        lGutschNrNeu = glGutschNr
    Else
        lGutschNrNeu = NewGutschein
    End If
    
    glGutschNr = lGutschNrNeu
    If lGutschNrNeu = lGutschNrMax Or lGutschNrNeu <= 0 Then
        MsgBox "Keine neue Gutschein-Nr erzeugt!", vbCritical, "STOP!"
        Exit Sub
    End If
    
    lbednu = Val(frmWKL20!Text1(0).Text)
    lDatum = Fix(Now)
    
    lKUNDNR = Val(frmWKL20!Label2(7).Caption)
    
    cSQL = "Select * from GUTSCH where GUTSCHNR = 0"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    rsrs.AddNew
    rsrs!gutschnr = lGutschNrNeu
    rsrs!BEDNU = lbednu
    rsrs!DAT_AUSG = lDatum
    rsrs!Wert = dWert
    rsrs!Kundnr = lKUNDNR
    rsrs!SYNStatus = "A"
    rsrs!Status = cWert
    rsrs!FILIALE = gcFilNr
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
    
    ProtokolliereALTGutscheinWK20g lKUNDNR, dWert, lGutschNrNeu
    
    
    cFeld = Trim$(Str$(lGutschNrNeu))
    cFeld = Space$(8 - Len(cFeld)) & cFeld
    cLBSatz = cFeld & " "
    
    cFeld = Format$(lDatum, "DD.MM.YYYY")
    cLBSatz = cLBSatz & cFeld & " "
    
    cFeld = Format$(dWert, "#####0.00")
    cFeld = Space$(10 - Len(cFeld)) & cFeld
    cLBSatz = cLBSatz & cFeld & " "
    
    cLBSatz = cLBSatz & "ALT"
    List3.AddItem cLBSatz
    
    AktualisiereZahlungWK20g cLBSatz, "+"
    
    'Dialog <Alter Gutschein> schließen
    SSCommand4_Click 13
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SchreibeAltenGutscheinWK20g"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
        
        Fehlermeldung1

    End If
End Sub


Private Sub Command5_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim ctmp As String
    Dim lcount As Long
    Dim iFehler As Integer
    
    Dim lCents      As Long
    Dim sCent       As String
    Dim lret1       As Long
    Dim sBLZ        As String * 2000
    Dim lBuffer     As Long

    Dim lerrCode    As Long
    Dim serrMeldung As String * 8000
    Dim lForcePIN   As Long
    
    iFehler = 0
    
    Command5(0).Enabled = False
    
    Select Case index
        Case Is = 0
            Screen.MousePointer = 11
            iFehler = 1
            
            Command5(0).Enabled = False ' Button aus
            Command5(1).Enabled = False 'Ok Button aus
            iFehler = 2
            If InStr(UCase$(Label15.Caption), "KARTENVERKAUF") > 0 Then
                iFehler = 3
                If gbEcash Then
                    Select Case gsEPartner

                        Case Is = "ADT"
                            iFehler = 5
                            If gsAdtVerfahren = "XML" Then

                                iFehler = 6
                                Dim hwnd&
                                Dim Y As String
                                Dim result&
                                Dim Title$
                                
                                Label28.Caption = "Bedienen Sie jetzt das Kartenterminal!"
                                Label28.Refresh
                                
                                Dim lRet        As Long
                                Dim iRet        As Integer
                                Dim sTraceNr    As String
                                iFehler = 7
                                If CDbl(Label6(5).Caption) < 0 Then
                                
                                    If CInt(gADTclientId) = 0 Then
                                        'Storno
                                        
                                        iFehler = 8
                                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
            
                                        iFehler = 9
                                        lRet = Shell("C:\Programme\EL-ME\SECpos\SECposPay\SECposPay.exe", vbHide) 'secpos
                                        AppActivate lRet
                                        iFehler = 10
                                        'erstmal zum storno navigieren
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        
                                        SendKeys "{Down}", True
                                        
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        SendKeys "{TAB}", True
                                        
                                        SendKeys "{Down}", True
                                        
                                        SendKeys sTraceNr, True
                                        SendKeys "{enter}", True
                                        
                                        iFehler = 11
                        
                                        Call keybd_event(VK_LWIN, 0, 0, 0)
                                        Call keybd_event(77, 0, 0, 0)
                                        Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
                                        iFehler = 12
                                        iRet = MsgBox("Storno ok? ", vbInformation + vbYesNo, "Winkiss Frage:")
                                        
                                        Y = "SECpos Pay" '  (Terminal-ID: " & gsTerminalid & ")"
                                        iFehler = 13
                                        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
                                    
                                        Do
                                            result = GetWindowTextLength(hwnd) + 1
                                            Title = Space(result)
                                            result = GetWindowText(hwnd, Title, result)
                                            Title = Left$(Title, Len(Title) - 1)
                                    
                                            If InStr(1, Title, Y) Then
                                    '            MsgBox hwnd
                                                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                                    
                                            End If
                                    
                                            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
                                        Loop Until hwnd = 0
                                        iFehler = 14
                                        
                                        If iRet = vbNo Then
                                            Screen.MousePointer = 0
                                            Command5(0).Enabled = True
                                            Command5(1).Enabled = True
                                            Exit Sub
                                        End If
                                        
                                        iFehler = 15
                                    Else
                                        'Storno neuer Weg
                                        If CInt(gADTclientId) > 0 Then
                                        
                                            If gADTipAdress <> "" And gADTport <> "" Then
                                                '192.168.1.14 '20002
                                                lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, gADTipAdress, gADTport, -1, -1, -1, vbNullString)
                                            Else
                                                lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, vbNullString, -1, -1, -1, -1, vbNullString)
                                            End If
                                            
                                        End If
                                        
                                        sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                                    
                                        lRet = ELMEReversal(CLng(sTraceNr))
  
                                        If lRet = 0 Then
                                            lRet = ELMEGetPrint(sBLZ, 8000)
                                            If lRet = 0 Then
'                                                MsgBox sBLZ
                                                gsAdtBeleg = sBLZ
                                            Else
                                                MsgBox "Fehler ELMEGetPrint: " & lRet, vbCritical, "Winkiss Fehler:"
                                                gsAdtBeleg = ""
                                            End If
                                        Else
'                                            MsgBox "Fehler ELMEPay: " & lRet
                                            lRet = ELMEGetLastError(lerrCode, serrMeldung, 8000)
                                            If lRet = 0 Then
                                                MsgBox serrMeldung, vbCritical, "Winkiss Fehler:"
                                            Else
                                                MsgBox "Fehler ELMEGetLastError: " & lRet, vbCritical, "Winkiss Fehler:"
                                            End If
                                            'Abbruch
                                            Screen.MousePointer = 0
                                            Command5(0).Enabled = True
                                            Command5(1).Enabled = True
                                            Exit Sub
                                            'Abbruch
                                        End If
                                    
                                    End If
                                Else
                                    'Zahlung
                                    
                                    If CInt(gADTclientId) = 0 Then
                                        iFehler = 16
                                        lRet = Shell("C:\Programme\EL-ME\SECpos\SECposPay\SECposPay.exe", vbHide) 'secpos
                                        AppActivate lRet
                                      
                                        iFehler = 17
                                        SendKeys Label6(5).Caption, True
                                        SendKeys "{enter}", True
                                        iFehler = 18
                                        Call keybd_event(VK_LWIN, 0, 0, 0)
                                        Call keybd_event(77, 0, 0, 0)
                                        Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
                        
                                        iFehler = 19
                                        iRet = MsgBox("Zahlung ok? ", vbInformation + vbYesNo, "Winkiss Frage:")
                                          
                                        Y = "SECpos Pay" '  (Terminal-ID: " & gsTerminalid & ")"
                                                
                                        hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
                                        iFehler = 20
                                        Do
                                            result = GetWindowTextLength(hwnd) + 1
                                            Title = Space(result)
                                            result = GetWindowText(hwnd, Title, result)
                                            Title = Left$(Title, Len(Title) - 1)
                                    
                                            If InStr(1, Title, Y) Then
                                                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                                            End If
                                    
                                            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
                                        Loop Until hwnd = 0
                                        iFehler = 21
                                          
                                        If iRet = vbNo Then
                                            Screen.MousePointer = 0
                                            Command5(0).Enabled = True
                                            Command5(1).Enabled = True
                                            Exit Sub
                                        End If
                                        iFehler = 22
                                    Else
                                        'Zahlung der neue Weg
                                        
                                        If CInt(gADTclientId) > 0 Then
                                            If gADTipAdress <> "" And gADTport <> "" Then
                                                '192.168.1.14 '20002
                                                lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, gADTipAdress, gADTport, -1, -1, -1, vbNullString)
                                            Else
                                                lRet = ELMESettings(vbNullString, gADTclientId, vbNullString, vbNullString, -1, -1, -1, -1, vbNullString)
                                            End If
                                        End If
                                        
                                        lForcePIN = 0
                                        sCent = Label6(5).Caption
                                        sCent = SwapStr(sCent, ",", "")
                                        lCents = CLng(sCent)
                                        lret1 = ELMEPay(lCents, lForcePIN)
    
                                        If lret1 = 0 Then

                                            lret1 = ELMEGetPrint(sBLZ, 2000)
                                            
                                            If lret1 = 0 Then

                                                gsAdtBeleg = sBLZ
                                            Else
                                                MsgBox "Fehler ELMEGetPrint: " & lret1, vbCritical, "Winkiss Fehler:"
                                                gsAdtBeleg = ""
                                            End If
                                        Else

                                            lret1 = ELMEGetLastError(lerrCode, serrMeldung, 8000)
                                            If lret1 = 0 Then
                                                MsgBox serrMeldung, vbCritical, "Winkiss Fehler:"
                                            Else
                                                MsgBox "Fehler ELMEGetLastError: " & lret1, vbCritical, "Winkiss Fehler:"
                                            End If
                                            'Abbruch
                                            Screen.MousePointer = 0
                                            Command5(0).Enabled = True
                                            Command5(1).Enabled = True
                                            Exit Sub
                                            'Abbruch
                                        End If
                                    End If
                                End If
                            ElseIf gsAdtVerfahren = "INOUT" Then
                            
                                iFehler = 23
                            End If
                        Case "ELP"
                        
                            Label28.Caption = "Bedienen Sie jetzt das Kartenterminal!"
                            Label28.Refresh
                                
                            If CDbl(Label6(5).Caption) < 0 Then
                                
                                'Storno
                                sCent = Label6(5).Caption
                                sCent = SwapStr(sCent, ",", "")
                                sCent = SwapStr(sCent, "-", "")
                                
                                sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "TA Nr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                                       
                                Storno_elPAY sTraceNr, sCent
                                
                                If giELPAY_Fehler > 0 Then
                                
                                    'Abbruch, so geht Abbruch
                                    Screen.MousePointer = 0
                                    Command5(0).Enabled = True
                                    Command5(1).Enabled = True
                                    
                                    Label28.Caption = "Fehler am Kartenterminal!"
                                    Label28.Refresh
                                    
                                    Exit Sub
                                    'Abbruch
                                End If
                                
                            Else
                
                                'Zahlung
                                sCent = Label6(5).Caption
                                sCent = SwapStr(sCent, ",", "")
                                
                                Zahlung_elPAY sCent
                                
                                If giELPAY_Fehler > 0 Then
                                
                                    'Abbruch, so geht Abbruch
                                    Screen.MousePointer = 0
                                    Command5(0).Enabled = True
                                    Command5(1).Enabled = True
                                    
                                    Label28.Caption = "Fehler am Kartenterminal!"
                                    Label28.Refresh
                                    
                                    Exit Sub
                                    'Abbruch
                                End If
                                    
                            End If
                            
                        Case "ZVT"
                        
                            Label28.Caption = "Bedienen Sie jetzt das Kartenterminal!"
                            Label28.Refresh
                                
                            If CDbl(Label6(5).Caption) < 0 Then
                                
                                'Storno
                                sCent = Label6(5).Caption
                                sCent = SwapStr(sCent, ",", "")
                                sCent = SwapStr(sCent, "-", "")
                                
                                dlgTaNr.Show 1
                    
                                sTraceNr = dlgTaNr.Back
'                                sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                                       
                                Storno_ZVT sTraceNr
                                
                                If giZVT_Fehler > 0 Then
                                
                                    'Abbruch, so geht Abbruch
                                    Screen.MousePointer = 0
                                    Command5(0).Enabled = True
                                    Command5(1).Enabled = True
                                    
                                    Label28.Caption = "Fehler am Kartenterminal!"
                                    Label28.Refresh
                                    
                                    Exit Sub
                                    'Abbruch
                                End If
                                
                            Else
                
                                'Zahlung
                                sCent = Label6(5).Caption
                                sCent = SwapStr(sCent, ",", "")
                                
                                Zahlung_ZVT sCent
                                
                                If giZVT_Fehler > 0 Then
                                
                                    'Abbruch, so geht Abbruch
                                    Screen.MousePointer = 0
                                    Command5(0).Enabled = True
                                    Command5(1).Enabled = True
                                    
                                    Label28.Caption = "Fehler am Kartenterminal!"
                                    Label28.Refresh
                                    
                                    Exit Sub
                                    'Abbruch
                                End If
                                    
                            End If
                            
                        Case "ZV2"
                        
                        
                            Label28.Caption = "Bedienen Sie jetzt das Kartenterminal!"
                            Label28.Refresh
                                
                            If CDbl(Label6(5).Caption) < 0 Then
                                
                                'Storno
                                sCent = Label6(5).Caption
                                sCent = SwapStr(sCent, ",", "")
                                sCent = SwapStr(sCent, "-", "")
                                
                                dlgTaNr.Show 1
                    
                                sTraceNr = dlgTaNr.Back
'                                sTraceNr = InputBox("Geben Sie bitte die" & vbCrLf & "BNr.:(steht auf dem Bon) ein!" & vbCrLf & "Bedienen Sie dann das Kartenterminal", "Winkiss Stornierung einer Kartenzahlung:")
                                       
                                Storno_ZVT2 sTraceNr, sCent, False
                                
                                If giZVT2_Fehler > 0 Then
                                
                                    'Abbruch, so geht Abbruch
                                    Screen.MousePointer = 0
                                    Command5(0).Enabled = True
                                    Command5(1).Enabled = True
                                    
                                    Label28.Caption = "Fehler am Kartenterminal!"
                                    Label28.Refresh
                                    
                                    Exit Sub
                                    'Abbruch
                                End If
                                
                            Else
                
                                'Zahlung
                                sCent = Label6(5).Caption
                                sCent = SwapStr(sCent, ",", "")
                                
                                Zahlung_ZVT2 sCent, False
                                
                                If giZVT2_Fehler > 0 Then
                                
                                    'Abbruch, so geht Abbruch
                                    Screen.MousePointer = 0
                                    Command5(0).Enabled = True
                                    Command5(1).Enabled = True
                                    
                                    Label28.Caption = "Fehler am Kartenterminal!"
                                    Label28.Refresh
                                    
                                    Exit Sub
                                    'Abbruch
                                End If
                                    
                            End If
                        
                        Case Else

                    End Select
                End If
                
                Label28.Caption = ""
                Label28.Refresh
                
                iFehler = 24
            End If
            
            Dim lKJADate As Long
            Dim cKJAZeit As String
            
            lKJADate = DateValue(Now)
            cKJAZeit = Format$(Now, "HH:MM:SS")
            
            
            
            
            
            
            
            
            Dim cErzielterPreis As String
            Dim cArtNr As String
            Dim dNichtUmsatz As Double
            dNichtUmsatz = 0
            Dim lAnzSatz As Long
            Dim lAktSatz As Long
            Dim cLBSatz As String
            Dim cUmsOK As String
            
            lAnzSatz = frmWKL20!List1.ListCount
            For lAktSatz = 0 To lAnzSatz - 1
                 cLBSatz = frmWKL20!List1.list(lAktSatz)
                
                If Len(cLBSatz) > 156 Then
                    cUmsOK = Mid(cLBSatz, 156, 1)
                Else
                    cUmsOK = "J"
                End If
        
                cArtNr = Mid(cLBSatz, 7, 6)
                
                cErzielterPreis = Mid(cLBSatz, 60, 9)
                cErzielterPreis = Trim$(cErzielterPreis)
                cErzielterPreis = fnMoveComma2Point$(cErzielterPreis)
                
                If cArtNr <> "666666" Then
                    If cUmsOK <> "N" Then
        '                dUmsatz = dUmsatz + Val(cErzielterPreis)
        '                dKundenZahl = 1
                    Else
                        dNichtUmsatz = dNichtUmsatz + Val(cErzielterPreis)
                    End If
                Else
        '            dWertGutschein = dWertGutschein + Val(cErzielterPreis)
                End If
            Next lAktSatz
            
            
            
            Dim dGeldwert As Double
            dGeldwert = CDbl(Label6(5).Caption)
            
            
            
            If InStr(UCase$(Label15.Caption), "KARTENVERKAUF") > 0 Then
                insertKKZAHLTE lKJADate, cKJAZeit, CStr(gdBonNr), gcKasNum, gcKreditKarte, dGeldwert
                
                If dGeldwert > 0 Then
                    If dNichtUmsatz > 0 Then
                    
                        If dNichtUmsatz > dGeldwert Then
                            eintragen_AFCSTAT_NUMSKARTE dGeldwert
                        Else
                            eintragen_AFCSTAT_NUMSKARTE dNichtUmsatz
                        End If
                        
                        dNichtUmsatz = dNichtUmsatz - dGeldwert
                    
                    End If
                End If
                
                
                
            End If
            
            
            
            ctmp = Label6(1).Caption
            ctmp = fnMoveComma2Point$(ctmp)
            gdSumme = Val(ctmp)
            gdGegeben = gdSumme
            gdZurueck = 0
        
            iFehler = 25
            
            
            Command5(0).Visible = False
            
            '***********Drucke Bon
            '*****

            Dim cFeld                   As String
            

            cFeld = Trim$(Label3(0).Caption)
            cFeld = fnMoveComma2Point$(cFeld)
            dZuZahlen = Val(cFeld)
            
            iFehler = 26
            
            cFeld = Trim$(Label3(1).Caption)
            cFeld = fnMoveComma2Point$(cFeld)
            dEingereichteGutscheine = Val(cFeld)
            
            
            iFehler = 27
            
            cFeld = Trim$(Label5(2).Caption)
            cFeld = fnMoveComma2Point$(cFeld)
            dNochOffen = Val(cFeld)
            
            iFehler = 28
            
            cFeld = Trim$(Text3.Text)
            cFeld = fnMoveComma2Point$(cFeld)
            dGegeben = Val(cFeld)
            
            iFehler = 29
            
            If InStr(cFeld, ".") = 0 Then
                dGegeben = dGegeben / 100
                cFeld = Format$(dGegeben, "#####0.00")
                Text3.Text = cFeld
            End If
            dWertRestGutschein = 0
            
            iFehler = 30
            gcZahlMittel = "KA"
            
            TSSBerechnung
            SendeDaten2DruckerGutschein2WK20g 0, dZuZahlen, dGegeben, dWertRestGutschein
            
            iFehler = 31
            
            dGegeben = 0
            iFehler = 32
            UpdateAFCStatGutscheinModul20 dZuZahlen, dEingereichteGutscheine, dNochOffen, dWertRestGutschein
            
            If dwertGutverkauf > 0 Then
            
                If dEingereichteGutscheine < dwertGutverkauf Then
                    dwertGutverkauf = dEingereichteGutscheine
                End If
                updateafcstat "GUTSCHGUTSCH", dwertGutverkauf, gcKasNum
                updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
            End If
            dwertGutverkauf = 0
            
            iFehler = 33
            
            'Todo TSE
            
            InsertAFCBuchGutscheinModul20 dEingereichteGutscheine
            
            If CheckofP = True Then
                InsertProvision
            End If
            If CheckofX = True Then
                InsertXMarkierung
            End If
            '***********Drucke Bon Ende
            
            If gbBONNEIN = False Then
'            If frmWKL20.Command1(12).Caption = "Bon JA!!!" Then
                iFehler = 34
                gbBonDruck = True
            Else
                iFehler = 35
                gbBonDruck = False
            End If
            
            iFehler = 36
            DoEvents
            gdBonNr = -1
            gcKreditKarte = ""
    
            iFehler = 37
            Screen.MousePointer = 0
            Command5(1).Enabled = True
            Command5(1).Caption = "Schließen"
            Command5(1).SetFocus
            
            iFehler = 38
            
            gBAgeschlossen = True 'Ok Kartenzahlung Fertig
            
            iFehler = 39
    
        Case Is = 1 'Schließen oder Zurück
    
            Screen.MousePointer = 11
   
            If Command5(0).Visible = False Then  'Schließen
                iFehler = 40
                frmWKL20.List1.Clear
                frmWKL20.List3.Nodes.Clear
                iFehler = 41
                frmWKL20.Label41(1).Caption = 0
                frmWKL20.Label2(6).Caption = "0,00"
                LeereDialogModul20
                
                Command5(0).Visible = True
                Command5(0).Enabled = True
                iFehler = 44
                Frame10.Visible = False
                If gbBEDLEER = False Then
                    gbNumTaste = True
                    frmWKL20.Text1(0).Text = gcBedienerNr
                Else
                    frmWKL20.Text1(0).Text = ""
                End If
                iFehler = 45
                giAndersZahlung = 0
                gcKreditKarte = ""
                gcZahlMittel = ""
                iFehler = 46
                If gbBEDLEER = True Then
                    If Command5(1).Caption = "Schließen" Then
                        frmWKL20.Text1(0).Text = ""
                        frmWKL20.Label1(8).Caption = ""
                    End If
                End If
                iFehler = 47
        
                If gsiGESRAB > 0 Then
                    frmWKL20.zeigeZwangsrabatt
                End If
                iFehler = 48
                
                If Not gbDisplay Then
                    InitKundenDisplayModul20
                End If
                iFehler = 49
                
                
                frmWKL20.Visible = True
                
                Unload frmWK20g
                iFehler = 50
                Screen.MousePointer = 0
    
            Else 'Zurück
                iFehler = 51
                
                gcKreditKarte = ""
                
                Command5(0).Enabled = True
                Frame10.Visible = False
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            
            
            
            
    
    End Select
    Screen.MousePointer = 0
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 5 Then
        Resume Next
    Else
        Command5(1).Enabled = True
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Command5_Click"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. " & iFehler
        
        Fehlermeldung1
        Resume Next
    End If
    
End Sub
Private Sub TSSBerechnung()
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzSatz    As Long
    Dim lAktSatz    As Long
    Dim cSQL        As String
    Dim ctmp        As String
    Dim cLBSatz     As String
    Dim cExtend     As String
    Dim cWasSuchteMan     As String
    Dim dLiNr       As Double
    Dim dEkpr       As Double
    Dim dWert       As Double
    Dim iFeld       As Integer
    Dim iDbNr       As Integer
    Dim rsrs        As Recordset
    Dim rsKJ        As Recordset
    Dim rsArt       As Recordset
    Dim cArtMWSt    As String
    
    '*** Kunden-Umsatz ****
    Dim dKdUmsatz   As Double
    Dim dKdBonus    As Double
    
    '*** KASSJOUR-Felder ****
    Dim cKJArtNr    As String
    Dim cKJBezeich  As String
    Dim cKJMenge    As String
    Dim cKJAZeit    As String
    Dim cKJKundNr   As String
    Dim cKJFiliale  As String
    Dim cKJKasNum   As String
    Dim cKJLiNr     As String
    Dim cKJLPZ      As String
    Dim cKJAGN      As String
    Dim cKJEAN      As String
    Dim cKJMwst     As String
    Dim cKJBelegNr  As String
    Dim cUmsOK      As String
    Dim cBonusOk    As String
    Dim cKJMopreis  As String
    
    Dim dKJEkpr     As Double
    Dim dKJVkpr     As Double
    Dim dKJPreis    As Double
    Dim dKJBest1    As Double
    Dim dVkPr       As Double
    Dim dKJPreis2   As Double
    Dim dSpanne     As Double
    Dim lKJADate    As Long
    Dim lKJBediener As Long
    Dim sArtnr      As String
    Dim IAbschluss  As Long
    Dim ierrz       As Integer
    Dim dGeldwert   As Double
    Dim sRechner    As String
    Dim sPreisKz    As String
    
    Dim lPos As Long
    
    Dim cpfaddb As String
        
    cpfaddb = gcDBPfad
    If Right$(cpfaddb, 1) <> "\" Then
        cpfaddb = cpfaddb & "\"
    End If
    
    
    ctmp = Trim$(frmWKL20!Label2(7).Caption)
    If Val(ctmp) < 0 Then
        ctmp = "0"
    End If
    cKJKundNr = ctmp
    
    If Val(cKJKundNr) > 0 Then
        'dann nach Preiskz fragen
        sPreisKz = ermPREISKZ(cKJKundNr)
    End If

    
    sRechner = rechnername
    ierrz = 0
    
    lAnzSatz = frmWKL20!List1.ListCount

    cSQL = "Delete from AFCB" & sRechner & " "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    For lAktSatz = 0 To lAnzSatz - 1

        iFeld = 1
        cKJArtNr = ""
        cKJBezeich = ""
        cKJMenge = ""
        dKJPreis = 0
        lKJADate = 0
        cKJAZeit = ""
        lKJBediener = 0
        cKJKundNr = ""
        cKJFiliale = ""
        cKJKasNum = ""
        cKJLiNr = ""
        cKJLPZ = ""
        cKJAGN = ""
        cKJEAN = ""
        cKJMwst = ""
        dKJEkpr = 0
        dKJVkpr = 0
        cKJBelegNr = ""
        dKJBest1 = 0
        
        
        cLBSatz = frmWKL20!List1.list(lAktSatz)
        
'        MsgBox cLBSatz

        'Besonderheiten am Satzende

        'hier Besonders Merkmal - wird in Mopreis kassjour gespeichert
        
        If Len(cLBSatz) > 175 Then
            cKJMopreis = Mid(cLBSatz, 177, 8)
        Else
            cKJMopreis = "0"
        End If

        If Len(cLBSatz) > 157 Then
            cExtend = Mid(cLBSatz, 158, 18)
        Else
            cExtend = ""
        End If
        
        
        
        
        ctmp = Mid(cLBSatz, 7, 6)
        ctmp = Trim$(ctmp)
        sArtnr = ctmp
        
        '***************************************************
        '* Zeile ZWISCHENSUMME darf nicht übernommen werden!
        '***************************************************
        
        If ctmp <> "000000" Then
            
            cSQL = "Select LPZ, AGN, EAN, EKPR, linr, MWST, UMS_OK, BONUS_OK, Spanne from Artikel where Artnr = " & ctmp
            Set rsArt = gdBase.OpenRecordset(cSQL)
            
'            FnOpenrecordset rsArt, cSQL, 1, gdBase

            If Not rsArt.EOF Then
                If Not IsNull(rsArt!LPZ) Then
                    cKJLPZ = rsArt!LPZ
                Else
                    cKJLPZ = ""
                End If
        
                If Not IsNull(rsArt!AGN) Then
                    cKJAGN = rsArt!AGN
                Else
                    cKJAGN = ""
                End If
                
                If Not IsNull(rsArt!EAN) Then
                    cKJEAN = rsArt!EAN
                Else
                    cKJEAN = ""
                End If
                
                If Not IsNull(rsArt!ekpr) Then
                    dEkpr = rsArt!ekpr
                Else
                    dEkpr = 0
                End If
                
                If Not IsNull(rsArt!linr) Then
                    dLiNr = rsArt!linr
                Else
                    dLiNr = 0
                End If
                
                If Not IsNull(rsArt!MWST) Then
                    cArtMWSt = rsArt!MWST
                Else
                    cArtMWSt = "V"
                End If
                
                
                'ist Preiskz = 6 also Netto dann mwst = O
                If Val(sPreisKz) = 6 Then
                    cArtMWSt = "O"
                End If
                
                
                If Not IsNull(rsArt!UMS_OK) Then
                    cUmsOK = rsArt!UMS_OK
                Else
                    cUmsOK = "J"
                End If
                
                If Not IsNull(rsArt!BONUS_OK) Then
                    cBonusOk = rsArt!BONUS_OK
                Else
                    cBonusOk = "J"
                End If
                
                If Not IsNull(rsArt!SPANNE) Then
                    dSpanne = rsArt!SPANNE
                Else
                    dSpanne = 0
                End If
                             
            Else
                dEkpr = 0
                dLiNr = 0
            End If
            
            rsArt.Close: Set rsArt = Nothing
            
            Set rsrs = gdBase.OpenRecordset("AFCB" & sRechner, dbOpenTable)
                        
            rsrs.AddNew
            rsrs!SYNStatus = "A"

            If ctmp = "666666" Then
                If gbGutscheinBeiVKversteuern = True Then
                    cBonusOk = "N"
                    cUmsOK = "J"
                    cArtMWSt = "V"
                Else
                    cBonusOk = "N"
                    cUmsOK = "N"
                    cArtMWSt = "O"
                End If
            End If

            ctmp = Mid(cLBSatz, 148, 3)
            ctmp = Trim$(ctmp)

            rsrs!abednu = Val(ctmp)
            lKJBediener = Val(ctmp)

            rsrs!AFLAG = 0
            
            If Left(cLBSatz, 1) = "x" Then
                ctmp = Mid(cLBSatz, 2, 4)
            Else
                ctmp = Mid(cLBSatz, 1, 5)
            End If

           
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!aMenge = Val(ctmp)
            cKJMenge = ctmp

            ctmp = Mid(cLBSatz, 60, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!APREIS = Val(ctmp)
            dKJPreis = rsrs!APREIS

            ctmp = Mid(cLBSatz, 7, 6)
            ctmp = Trim$(ctmp)
            rsrs!aartnr = Val(ctmp)
            cKJArtNr = ctmp

            ctmp = Mid(cLBSatz, 14, 35)
            ctmp = Trim$(ctmp)

            rsrs!ABEZEICH = ctmp
            cKJBezeich = ctmp

            rsrs!ADATE = Fix(Now)
            rsrs!AZEIT = Format$(Now, "HH:MM:SS")
            lKJADate = rsrs!ADATE
            cKJAZeit = rsrs!AZEIT

            rsrs!AMWSK = cArtMWSt
            cKJMwst = cArtMWSt

            If ctmp = "V" Then
                ctmp = Mid(cLBSatz, 104, 9)
            ElseIf ctmp = "E" Then
                ctmp = Mid(cLBSatz, 114, 9)
            Else
                ctmp = "0"
            End If
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            rsrs!AMWST = Val(ctmp)

            ctmp = frmWKL20!Label2(7).Caption
            ctmp = Trim$(ctmp)
            If Val(ctmp) < 0 Then
                ctmp = "0"
            End If
            rsrs!AKUNUM = Val(ctmp)
            cKJKundNr = ctmp

            rsrs!BELEGNR = gdBonNr
            cKJBelegNr = rsrs!BELEGNR

            rsrs!kasnum = Val(gcKasNum)
            cKJFiliale = gcFilNr

            rsrs!BUCHFLAG = 0

            ctmp = Mid(cLBSatz, 50, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)
            
            rsrs!AALTPREIS = Format(Val(ctmp), "#####0.00")

            ctmp = Mid(cLBSatz, 128, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            If Val(ctmp) = 0 Then
                ctmp = Mid(cLBSatz, 50, 9)
                ctmp = Trim$(ctmp)
                ctmp = fnMoveComma2Point$(ctmp)
                
                rsrs!AVKPR = Format(Val(ctmp), "#####0.00")

                dKJVkpr = rsrs!AVKPR
            Else
                rsrs!AVKPR = Val(ctmp)
                dKJVkpr = rsrs!AVKPR
            End If

            If dEkpr = 0 Then
            
                If sArtnr = "666668" Or sArtnr = "666669" Then
                    If gdZeitungsSpanne <> 0 Then
                        dEkpr = EKausNettospanneerrechnen(gdZeitungsSpanne, Val(ctmp), cArtMWSt)
                    End If
                Else
                    If dSpanne <> 0 Then
                        dEkpr = EKausNettospanneerrechnen(dSpanne, Val(ctmp), cArtMWSt)
                    End If
                End If
            
            End If

            rsrs!ALEKPR = dEkpr
            dKJEkpr = dEkpr

            rsrs!linr = dLiNr
            cKJLiNr = Trim$(Str$(dLiNr))

            If gcKreditKarte <> "" Then
                rsrs!kk_art = gcKreditKarte
            Else
                rsrs!kk_art = gcZahlMittel
            End If

            ctmp = Mid(cLBSatz, 138, 9)
            ctmp = Trim$(ctmp)
            ctmp = fnMoveComma2Point$(ctmp)

            dKJBest1 = Val(ctmp)
            rsrs!BESTAND = Val(ctmp)
            rsrs!ZHLGGUTSCH = 0
            

            
            rsrs!UMS_OK = cUmsOK
            rsrs!BONUS_OK = cBonusOk
            rsrs!FILIALNR = Val(gcFilNr)
            rsrs.Update

            rsrs.Close: Set rsrs = Nothing

        End If
    Next lAktSatz
    
    
'    If gcZahlMittel = "BA" Then
'        If gbGutscheinBeiVKversteuern = True Then
'            cSQL = "Select sum(preis) as NICHTUMS from KJ" & sRechner & " where ums_OK = 'N' and kk_art = 'BA'  "
'        Else
'            cSQL = "Select sum(preis) as NICHTUMS from KJ" & sRechner & " where ums_OK = 'N' and kk_art = 'BA' and ARTNR <> 666666 "
'        End If
'
'        Set rsRS = gdBase.OpenRecordset(cSQL)
'        If Not rsRS.EOF Then
'            If Not IsNull(rsRS!NICHTUMS) Then
'                insertNichtUmsBar lKJADate, cKJAZeit, cKJBelegNr, gcKasNum, CDbl(rsRS!NICHTUMS)
'            End If
'        End If
'    End If
'
    
    
    '''''''''Oliver Alte TSE START
    
'    'TODO TSE FINISH
'    Dim dUmsatzVolleMwst As Double
'    Dim dUmsatzErmMwst As Double
'    Dim dUmsatzOhneMwst As Double
'
'    Dim dUmsatzGesamt As Double
'
'
'
'    dUmsatzVolleMwst = ermWertforTSS("V", "J", "AFCB" & sRechner)
'    dUmsatzErmMwst = ermWertforTSS("E", "J", "AFCB" & sRechner)
'    dUmsatzOhneMwst = ermWertforTSS("O", "J", "AFCB" & sRechner)
'
'    dUmsatzGesamt = dUmsatzVolleMwst + dUmsatzErmMwst + dUmsatzOhneMwst
'
'    If gcZahlMittel = "BA" Then
'        TSS.Finish frmWKL20!WinsockTSE, Beleg, dUmsatzVolleMwst, dUmsatzErmMwst, dUmsatzOhneMwst, "EUR", dUmsatzGesamt, 0
'    Else
'        TSS.Finish frmWKL20!WinsockTSE, Beleg, dUmsatzVolleMwst, dUmsatzErmMwst, dUmsatzOhneMwst, "EUR", 0, dUmsatzGesamt
'    End If
'
    '''''''''Oliver Alte TSE ENDE
    
   

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "TSSBerechnung"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Private Sub Form_Activate()
On Error GoTo LOKAL_ERROR

    Text1.SetFocus
    Me.Refresh
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 5 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "Form_Activate"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
        
        Fehlermeldung1
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    WK20gPositionieren

    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, Nothing
    
    sSort = "asc"
    
    grunde
    
    dwertGutverkauf = 0
    dwertGutverkauf = GutscheinDabei
    
    bFirstEingabe = False
    
    If gbAlterGutschein_Ausblenden = True Then
        SSCommand2(1).Enabled = False
    End If
    
    Me.Refresh
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1

End Sub
Private Sub grunde()
    On Error GoTo LOKAL_ERROR
    
    Frame8.Visible = False
    Frame1.Visible = True
    Frame2.Visible = True
    Frame3.Visible = True
    Frame5.Visible = True
    
    gBAgeschlossen = False
    gbGutschein = False
    bseekerfolg = False
    
    'Basiswährung setzen
    Label2(3).Caption = gcWaehrung
    Label2(4).Caption = gcWaehrung
    Label2(5).Caption = gcWaehrung

    'Anzeigelisten leeren
    List1.Clear     'Liste für Überschriften
    List2.Clear     'Liste mit offenen Gutscheinen
    List3.Clear     'Liste mit gewählten Gutscheinen

    'offene Gutscheine einlesen
    LeseOffeneGutscheineWK20g "gutschnr" & " " & sSort

    'zu zahlenden Betrag in Dialog bringen
    Label3(0).Caption = frmWKL20!Label2(6).Caption
    Label3(2).Caption = frmWKL20!Label2(6).Caption
    Label3(1).Caption = "0,00"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "grunde"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
End Sub
Private Sub LeseOffeneGutscheineWK20g(sOrder As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cFeld       As String
    Dim cLBSatz     As String
    Dim dWert       As Double
    Dim rsrs        As Recordset
    Dim lcount      As Long
    Dim lcountall   As Long
    Dim cGutschnr   As String
    Dim j As Integer
    
    List1.Clear
    List2.Clear
    
    List1.AddItem "Gutsch. Ausgabe am       Wert"
    
    sSQL = "Select * from gutsch where Status <> 'L' and not Wert  is null order by " & sOrder & ""  'gutschnr"

    frmWKL20.picprogress.Visible = True

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        lcountall = lcount
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            lcount = lcount - 1
            
            j = lcount Mod 3000
            If j = 0 Then
                frmWKL20.txtStatus.Text = CStr(lcount * 100 / lcountall)
            Else
                
            End If
            
            If IsNull(rsrs!DAT_EINL) Or rsrs!DAT_EINL = 0 Then
                If Not IsNull(rsrs!gutschnr) Then
                    cFeld = rsrs!gutschnr
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cGutschnr = cFeld
                cFeld = Space$(8 - Len(cFeld)) & cFeld
                cLBSatz = cFeld & " "
                
                If Not IsNull(rsrs!DAT_AUSG) Then
                    dWert = rsrs!DAT_AUSG
                Else
                    dWert = 0
                End If
                If dWert > 0 Then
                    cFeld = Format$(dWert, "DD.MM.YYYY")
                Else
                    cFeld = ""
                End If
                cFeld = Trim$(cFeld)
                cFeld = cFeld & Space$(10 - Len(cFeld))
                cLBSatz = cLBSatz & cFeld & " "
                
                If Not IsNull(rsrs!Wert) Then
                    dWert = rsrs!Wert
                Else
                    dWert = 0
                End If
                cFeld = Format$(dWert, "######0.00")
                cFeld = Trim$(cFeld)
                cFeld = Space$(10 - Len(cFeld)) & cFeld
                cLBSatz = cLBSatz & cFeld & " "
                
                List2.AddItem cLBSatz
                
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    frmWKL20.picprogress.Visible = False
    
    Label6(1).Caption = "vorhandene Gutscheine: " & lcountall
    Label6(1).Refresh
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 5 Then
        sSQL = "Update gutsch set Status = 'L' where gutschnr = " & cGutschnr
        gdBase.Execute sSQL, dbFailOnError
        
        cGutschnr = ""
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "LeseOffeneGutscheineWK20g"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
        
        Fehlermeldung1
    End If
'    Resume Next
End Sub

Private Sub SSCommand1_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
        
    Dim lcount As Long
    Dim cLBSatz As String
    Dim cFeld As String
    Dim dSuchWert As Double
    Dim dFindeWert As Double
    
    bseekerfolg = False
    Select Case index
        Case Is = 0     'Suche Gutschein nach Wert
            If Trim$(Text1.Text) = "" Then
            
                If sSort = "asc" Then
                    sSort = "desc"
                Else
                    sSort = "asc"
                End If
                LeseOffeneGutscheineWK20g "WERT" & " " & sSort
                
'                MsgBox "Bitte den Wert des gesuchten Gutscheines eingeben!", vbCritical, "STOP!"
                Text1.SetFocus
            Else
                cFeld = Text1.Text
                cFeld = fnMoveComma2Point$(cFeld)
                dSuchWert = Val(cFeld)
                If InStr(cFeld, ".") = 0 Then
                    dSuchWert = dSuchWert / 100
                    cFeld = Format$(dSuchWert, "#####0.00")
                    Text1.Text = cFeld
                End If
                For lcount = 0 To List2.ListCount - 1
                    cLBSatz = List2.list(lcount)
                    cFeld = Mid(cLBSatz, 22, 9)
                    cFeld = fnMoveComma2Point$(cFeld)
                    dFindeWert = Val(cFeld)
                    If dFindeWert = dSuchWert Then
                        List2.Selected(lcount) = True
                        Exit For
                    End If
                Next lcount
            End If
        Case Is = 1     'Suche Gutschein nach Nr
            If Trim$(Text1.Text) = "" Then
            
                If sSort = "asc" Then
                    sSort = "desc"
                Else
                    sSort = "asc"
                End If
                LeseOffeneGutscheineWK20g "GUTSCHNR" & " " & sSort
            
                
'                MsgBox "Bitte die Nummer des gesuchten Gutscheines eingeben!", vbCritical, "STOP!"
                Text1.SetFocus
            Else
            
                cFeld = Text1.Text
                cFeld = fnMoveComma2Point$(cFeld)
                
                'selbst erstellter
                If Len(cFeld) = 8 Then
                    If Left(cFeld, 1) = "2" Then
                        cFeld = Mid(cFeld, 2, 6)
                        
                    ElseIf Left(cFeld, 1) = "0" Then
                        cFeld = Mid(cFeld, 2, 6)
                        
                    ElseIf Left(cFeld, 1) = "9" Then
                        cFeld = Mid(cFeld, 2, 6)
                    
                    End If
                End If
                
                'Gottmann Goedecke
                If Left(cFeld, 2) = "00" Then
                    cFeld = Left(cFeld, Len(cFeld) - 1)
                End If
                
                
                'selbst erstellter 13er an der Kasse
                If Len(cFeld) = 13 Then
                
                    If Left(cFeld, 1) = "2" And gbGutschnrKomplett = True Then
                        cFeld = Left(cFeld, 8)
                    End If
                
                
                    If Left(cFeld, 2) = "22" Or Left(cFeld, 2) = "21" Then
                        cFeld = Mid(cFeld, 3, 10)

                    End If
                End If
                
                dSuchWert = Val(cFeld)
                
                For lcount = 0 To List2.ListCount '- 1
                    cLBSatz = List2.list(lcount)
                    cFeld = Left(cLBSatz, 8)
                    cFeld = fnMoveComma2Point$(cFeld)
                    dFindeWert = Val(cFeld)
                    If dFindeWert = dSuchWert Then
                        List2.Selected(lcount) = True
                        bseekerfolg = True
                        Exit For
                    End If
                    
                Next lcount
                If bseekerfolg = False Then
                    Text1.Text = ""
                    Text1.SetFocus
                    MsgBox "Kein Gutschein gefunden", vbInformation, "Winkiss Hinweis:"
                Else
                    Text1.Text = ""
                    Text1.SetFocus
                End If
                    
            End If
        
        Case 2 To 11    'Ziffern
            Text1.Text = Text1.Text & SSCommand1(index).Caption
            Text1.SetFocus
            
        Case Is = 12    'Komma
            If InStr(Text1.Text, ",") = 0 Then
                Text1.Text = Text1.Text & SSCommand1(index).Caption
            End If
            Text1.SetFocus
            
        Case Is = 13    'Clear
            Text1.Text = ""
            Text1.SetFocus
            
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Function GutscheinDabei() As Double
On Error GoTo LOKAL_ERROR

Dim cLBSatz As String
Dim dAPreis As Double
Dim dGAPreis As Double
Dim cArtNr As String
Dim lAnzSatz As Long
Dim lAktSatz As Long
Dim ctmp As String

GutscheinDabei = 0
dGAPreis = 0
lAnzSatz = frmWKL20!List1.ListCount
For lAktSatz = 0 To lAnzSatz - 1
    cLBSatz = frmWKL20!List1.list(lAktSatz)

    cArtNr = Mid(cLBSatz, 7, 6)
    cArtNr = Trim$(cArtNr)
    
    If Val(cArtNr) = 666666 Then
        'Yes = Preis ermitteln
        dAPreis = 0
        ctmp = Mid(cLBSatz, 60, 9)
        ctmp = Trim$(ctmp)
        dAPreis = Val(ctmp)
        dGAPreis = dGAPreis + dAPreis
    End If
Next lAktSatz

GutscheinDabei = dGAPreis
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GutscheinDabei"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
End Function
Private Sub SSCommand2_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    Dim cFeld As String
    Dim dWert As Double
    
    
    Dim lGutschnr As Long
    Dim dRestGutschein As Double
    Dim cNotiz          As String
        
    cFeld = Trim$(Label3(0).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dZuZahlen = Val(cFeld)
    
    cFeld = Trim$(Label3(1).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dEingereichteGutscheine = Val(cFeld)

    dRestGutschein = 0
    
    Select Case index
        Case Is = 0     'Gutschein auswählen
            If List2.ListIndex < 0 Then
                MsgBox "Bitte einen Gutschein auswählen!", vbInformation, "Winkiss Hinweis:"
                List2.SetFocus
                Exit Sub
            End If
            
            cLBSatz = List2.list(List2.ListIndex)
            
            cNotiz = ermittleGutschNotizen(Trim(Left(cLBSatz, 8)))
            If cNotiz <> "" Then
                Text4.Visible = True
                Text4.Text = cNotiz
            Else
                Text4.Visible = False
                Text4.Text = ""
            End If
            
            AktualisiereZahlungWK20g cLBSatz, "+"
            
            List3.AddItem cLBSatz
            List2.RemoveItem List2.ListIndex
            
        Case Is = 1     'alter Gutschein
            Text2.Text = ""
            Frame1.Visible = False
            Frame2.Visible = False
            Frame3.Visible = False
            Frame5.Visible = False
            Frame6.Visible = True
            
        Case Is = 2     'Abbrechen
            If List3.ListCount > 0 Then
                cFeld = "Verlassen des Dialoges nicht möglich," & vbCrLf
                cFeld = cFeld & "da noch Gutscheine zur Zahlung ausgewählt sind!"
                
                MsgBox cFeld, vbCritical, "STOP!"
                List3.SetFocus
            Else
                gbBackaus20g = True
                Unload frmWK20g
            End If
            
        Case Is = 3     'Kassieren
            If List3.ListCount = 0 Then
                MsgBox "Kassieren nicht möglich! Kein Gutschein ausgewählt!", vbInformation, "Winkiss Hinweis:"
                Exit Sub
            End If
            cFeld = Label3(2).Caption
            cFeld = fnMoveComma2Point$(cFeld)
            dWert = Val(cFeld)
            
            
            
            Select Case dWert
                Case Is = 0     'Bezahlung paßt!
                    gcRueckgeld = "0,00"
                    dGegeben = 0
                    dNochOffen = 0
                    lGutschnr = 0
                    dRestGutschein = 0
                    
                    TSSBerechnung
                    SendeDaten2DruckerGutschein2WK20g lGutschnr, dZuZahlen, dGegeben, dRestGutschein
                    UpdateAFCStatGutscheinModul20 dZuZahlen, dEingereichteGutscheine, dNochOffen, dRestGutschein
                    InsertAFCBuchGutscheinModul20 dEingereichteGutscheine
                    If CheckofP = True Then
                        InsertProvision
                    End If
                    If CheckofX = True Then
                        InsertXMarkierung
                    End If
                    
                    If dwertGutverkauf > 0 Then
                        updateafcstat "GUTSCHGUTSCH", dwertGutverkauf, gcKasNum
                        updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
                    End If
                    dwertGutverkauf = 0
        
                    frmWK20a.Show 1
                    ReInitDialog20WK20g
                    Unload frmWK20g
                    
                Case Is < 0     'Überzahlung: Restgutschein oder Auszahlung
                    Label5(0).Caption = Format$(Abs(dWert), "#####0.00")
                    Label5(1).Caption = Format$(Abs(dWert), "#####0.00")
                    
'                    MsgBox Format$(Abs(dWert), "#####0.00")
                    
                    If Format$(Abs(dWert), "#####0.00") <= gdRESTGU Then
                        Text5.Text = Format$(Abs(dWert), "#####0.00")
                    Else
                        Text5.Text = ""
                    End If
                    
                    Label4(4).Caption = gcWaehrung
                    Label4(5).Caption = gcWaehrung
                    Label4(6).Caption = gcWaehrung
                    Frame1.Visible = False
                    Frame2.Visible = False
                    Frame3.Visible = False
                    Frame5.Visible = False
                    Frame7.Visible = True
                    
                Case Is > 0     'Unterdeckung: Restzahlung durch BAR, SCHECK, KARTE
                    Label5(2).Caption = Format$(dWert, "#####0.00")
                    Text3.Text = Format$(dWert, "#####0.00")
                    Label4(8).Caption = gcWaehrung
                    Label4(10).Caption = gcWaehrung
                    Frame1.Visible = False
                    Frame2.Visible = False
                    Frame3.Visible = False
                    Frame5.Visible = False
                    Frame8.Visible = True
                    SSCommand8(0).Enabled = True
                    If gbEcash Then
                        SSCommand8(3).Enabled = False
                    Else
                        SSCommand8(3).Enabled = True
                    End If
            End Select
    End Select
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub SSCommand3_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lRet        As Long
    Dim cLBSatz     As String
    Dim cFeld       As String
    
    If List3.ListIndex < 0 Then
        MsgBox "Bitte einen Gutschein in der Liste auswählen!", vbCritical, "STOP!"
        List3.SetFocus
        Exit Sub
    End If
    
    cLBSatz = List3.list(List3.ListIndex)
    Select Case index
        Case Is = 0     '** Gutschein-Auswahl zurücknehmen **
            If InStr(cLBSatz, "ALT") > 0 Then
                cFeld = "Sie wollen einen frisch erfaßten Alt-Gutschein zurücknehmen." & vbCrLf
                cFeld = cFeld & "Damit wird der Alt-Gutschein dauerhaft gültig." & vbCrLf & vbCrLf
                cFeld = cFeld & "Wollen Sie den Alt-Gutschein wirklich in den Gutscheinbestand übernehmen?"
                
                lRet = MsgBox(cFeld, vbQuestion + vbYesNo, "ALT-GUTSCHEIN")
                If lRet <> vbYes Then
                    Exit Sub
                End If
                cLBSatz = Left(cLBSatz, Len(cLBSatz) - 4)
            End If
            AktualisiereZahlungWK20g cLBSatz, "-"
            List2.AddItem cLBSatz
            List3.RemoveItem List3.ListIndex
            
        Case Is = 1     '** frisch erfaßten Alt-Gutschein wieder löschen **
        
            If InStr(cLBSatz, "ALT") = 0 Then
                cFeld = "Löschen ist nur bei neu erfaßten Alt-Gutscheinen möglich!" & vbCrLf & vbCrLf
                cFeld = cFeld & "Löschen von Gutscheinen siehe unter:" & vbCrLf
                cFeld = cFeld & "<LISTEN>" & vbCrLf
                cFeld = cFeld & "     -> <ARTIKELLISTE> " & vbCrLf
                cFeld = cFeld & "          -> <GUTSCHEINE>" & vbCrLf
                MsgBox cFeld, vbCritical, "STOP!"
                Exit Sub
            End If
            lRet = MsgBox("Wollen Sie den Gutschein " & vbCrLf & vbCrLf & cLBSatz & vbCrLf & vbCrLf & " wirklich löschen?", vbQuestion + vbYesNo, "LÖSCHEN")
            If lRet = vbYes Then
                LoescheAltenGutscheinWK20g cLBSatz
                AktualisiereZahlungWK20g cLBSatz, "-"
                List3.RemoveItem List3.ListIndex
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand3_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub SSCommand4_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
        
    Dim lcount          As Long
    Dim cLBSatz         As String
    Dim cFeld           As String
    Dim dSuchWert       As Double
    Dim dFindeWert      As Double
        
    Select Case index
        Case 0 To 9    'Ziffern
            Text2.Text = Text2.Text & SSCommand4(index).Caption
            Text2.SetFocus
            
        Case Is = 10    'Komma
            If InStr(Text2.Text, ",") = 0 Then
                Text2.Text = Text2.Text & SSCommand4(index).Caption
            End If
            Text2.SetFocus
            
        Case Is = 11    'Clear
            Text2.Text = ""
            Text2.SetFocus
            
        Case Is = 12    'Speichern
            SchreibeAltenGutscheinWK20g
        
        Case Is = 13    'Abbrechen
            Frame1.Visible = True
            Frame2.Visible = True
            Frame3.Visible = True
            Frame5.Visible = True
            Frame6.Visible = False
        
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand4_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub SSCommand5_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lRet                    As Long
    Dim lGutschnr               As Long
    Dim cFeld                   As String
    Dim dGeldAnKunden           As Double
    Dim dGeldAnKundenInBar      As Double
    Dim lbednu                  As Long
    Dim lDatum                  As Long
    Dim cSQL                    As String
    Dim rsrs                    As Recordset

    cFeld = Trim$(Label3(0).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dZuZahlen = Val(cFeld)
    
    cFeld = Trim$(Label3(1).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dEingereichteGutscheine = Val(cFeld)
    
    cFeld = Trim$(Label5(0).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dGeldAnKunden = Val(cFeld)
    
    cFeld = Trim$(Text5.Text)
    cFeld = fnMoveComma2Point$(cFeld)
    dGeldAnKundenInBar = Val(cFeld)
    If InStr(cFeld, ".") = 0 Then
        dGeldAnKundenInBar = dGeldAnKundenInBar / 100
        cFeld = Format$(dGeldAnKundenInBar, "#####0.00")
        Text5.Text = cFeld
    End If
    
    Select Case index
        Case Is = 0     '** OK **
            
            dNochOffen = dGeldAnKunden - dGeldAnKundenInBar
            If dGeldAnKundenInBar > dGeldAnKunden Then
                MsgBox "Zu hohe Bar-Rückzahlung!", vbCritical, "STOP!"
                Text5.SetFocus
                Exit Sub
            End If
            dWertRestGutschein = dGeldAnKunden - dGeldAnKundenInBar
            Label5(1).Caption = Format(dWertRestGutschein, "######0.00")
            If dWertRestGutschein > 0 Then
                If gbRGO = False Then
                    lGutschnr = NewGutschein
                    
                    If lGutschnr = 0 Then
                        Exit Sub
                    End If
                    
                    lbednu = Val(frmWKL20!Text1(0).Text)
                    lDatum = Fix(Now)
                    
                    cSQL = "Select * from GUTSCH where GUTSCHNR = 0"
                    Set rsrs = gdBase.OpenRecordset(cSQL)
                    rsrs.AddNew
                    rsrs!gutschnr = lGutschnr
                    rsrs!BEDNU = lbednu
                    rsrs!DAT_AUSG = lDatum
                    rsrs!Wert = dWertRestGutschein
                    rsrs!SYNStatus = "A"
                    rsrs!Status = "A"
                    rsrs!FILIALE = gcFilNr
                    rsrs.Update
                    rsrs.Close: Set rsrs = Nothing
                Else
                
                    If List3.ListCount < 0 Then
                        MsgBox "Es ist kein Gutschein ausgewählt.", vbInformation, "Winkiss Hinweis:"
                        Exit Sub
                    Else
                        'such den firstgutschein
                    End If
            
                    'Restgutschein = Originalgutschein
                    
                    'Achtung Achtung nicht den ersten Gutschein nehmen, sondern den mit dem höchsten Wert
'                    lGutschNr = CLng(Trim(Left(List3.list(0), 8)))

                    Dim cGutschnr As String
                    Dim cGutschWert As String
                    Dim dGutschwert As Double
                    Dim dMaxWert As Double
                    Dim cMaxGutschnr As String
                    Dim lAktSatz As Long
                    Dim cLBSatz As String
                    dMaxWert = 0

                    For lAktSatz = 0 To List3.ListCount - 1
                        cLBSatz = List3.list(lAktSatz)
                        cGutschnr = Trim(Left(cLBSatz, 8))
                       
                        cGutschWert = Trim(Mid(cLBSatz, 21, 10))
                        dGutschwert = CDbl(cGutschWert)
                        If dMaxWert < dGutschwert Then
                            dMaxWert = dGutschwert
                            cMaxGutschnr = cGutschnr
                        End If
                    Next lAktSatz

                    lGutschnr = CLng(cMaxGutschnr)

                    lbednu = Val(frmWKL20!Text1(0).Text)
                    lDatum = Fix(Now)
                    
                    cSQL = "Select * from GUTSCH where GUTSCHNR = " & lGutschnr
                    Set rsrs = gdBase.OpenRecordset(cSQL)
                    If Not rsrs.EOF Then
                        rsrs.Edit
                        rsrs!gutschnr = lGutschnr
                        rsrs!BEDNU = lbednu
                        rsrs!DAT_AUSG = lDatum
                        rsrs!DAT_EINL = Null
                        rsrs!Wert = dWertRestGutschein
                        rsrs!SYNStatus = "E"
                        rsrs!Status = "E"
                        rsrs!FILIALE = gcFilNr
                        rsrs.Update
                        rsrs.Close: Set rsrs = Nothing
                    End If
                
                End If
                ProtokolliereRueckGutscheinWK20g dWertRestGutschein, lGutschnr
            End If
            
            SSCommand5(0).Enabled = False
            dNochOffen = 0
            gcRueckgeld = Format$(dGeldAnKundenInBar, "######0.00")
            '2
            
            TSSBerechnung
            SendeDaten2DruckerGutschein2WK20g lGutschnr, dZuZahlen, dGegeben, dWertRestGutschein
            UpdateAFCStatGutscheinModul20 dZuZahlen, dEingereichteGutscheine, dNochOffen, dWertRestGutschein
            InsertAFCBuchGutscheinModul20 dEingereichteGutscheine
            If CheckofP = True Then
                InsertProvision
            End If
            
            If CheckofX = True Then
                InsertXMarkierung
            End If
            
            If dwertGutverkauf > 0 Then
                updateafcstat "GUTSCHGUTSCH", dwertGutverkauf, gcKasNum
                updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
            End If
            dwertGutverkauf = 0
            
            frmWK20a.Show 1
            ReInitDialog20WK20g
            Unload frmWK20g
            
            If gsiGESRAB > 0 Then
                frmWKL20.zeigeZwangsrabatt
            End If
            
        Case Is = 1     '** Abbrechen **
            Frame1.Visible = True
            Frame2.Visible = True
            Frame3.Visible = True
            Frame5.Visible = True
            Frame7.Visible = False
            bFirstEingabe = False
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand5_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub SSCommand6_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
        
        
    If bFirstEingabe = False Then
        Text5.Text = ""
        bFirstEingabe = True
    End If
        
        
        
    Select Case index
        Case 0 To 9    '** Ziffern **
        
        
        
            
            
            Text5.Text = Text5.Text & SSCommand6(index).Caption
            Text5.SetFocus
            
        Case Is = 10    '** Komma **
            If InStr(Text5.Text, ",") = 0 Then
                Text5.Text = Text5.Text & SSCommand6(index).Caption
            End If
            Text5.SetFocus
            
        Case Is = 11    '** Clear **
            Text5.Text = ""
            Text5.SetFocus
            
        Case Is = 12    '** Komma **
            If InStr(Text5.Text, "-") = 0 Then
                Text5.Text = SSCommand6(index).Caption & Text5.Text
            End If
            Text5.SetFocus
            
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand6_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1

End Sub
Private Sub SSCommand7_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    
    If bFirstEingabe = False Then
        Text3.Text = ""
        bFirstEingabe = True
    End If
        
    Select Case index
        Case 0 To 9    '** Ziffern **
            Text3.Text = Text3.Text & SSCommand7(index).Caption
            Text3.SetFocus
            
        Case Is = 10    '** Komma **
            If InStr(Text3.Text, ",") = 0 Then
                Text3.Text = Text3.Text & SSCommand7(index).Caption
            End If
            Text3.SetFocus
            
        Case Is = 11    '** Clear **
            Text3.Text = ""
            Text3.SetFocus
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand7_Click"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub SSCommand8_Click(index As Integer)
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim lRet                    As Long
    Dim cFeld                   As String
    Dim ctmp                    As String
    Dim cZeile2                 As String
    Dim cZeile1                 As String
    Dim sNochOffen              As String
    Dim dWert                   As Double
    Dim dGeldAnKunden           As Double
    Dim ifehl                   As Integer
    
    gbGutschOverBar = False
    
    ifehl = 1
    
    cFeld = Trim$(Label3(0).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dZuZahlen = Val(cFeld)
    
    ifehl = 2
    
    cFeld = Trim$(Label3(1).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dEingereichteGutscheine = Val(cFeld)
    
    ifehl = 3
    
    cFeld = Trim$(Label5(2).Caption)
    cFeld = fnMoveComma2Point$(cFeld)
    dNochOffen = Val(cFeld)
    
    ifehl = 4
    
    cFeld = Trim$(Text3.Text)
    cFeld = fnMoveComma2Point$(cFeld)
    dGegeben = Val(cFeld)
    If InStr(cFeld, ".") = 0 Then
        dGegeben = dGegeben / 100
        cFeld = Format$(dGegeben, "#####0.00")
        Text3.Text = cFeld
    End If
    dWertRestGutschein = 0
    
    ifehl = 5
    
    If index < 4 Then
        If index = 3 Then
            If dGegeben > dNochOffen Then
                dGeldAnKunden = dGegeben - dNochOffen
            Else
                dGeldAnKunden = 0
            End If
        Else
            If dGegeben < dNochOffen Then
                MsgBox "Die Restzahlung reicht nicht aus!", vbInformation, "Winkiss Hinweis:"
                Text3.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            ElseIf dGegeben > dNochOffen Then
                dGeldAnKunden = dGegeben - dNochOffen
            Else
                dGeldAnKunden = 0
            End If

        End If
    End If
    
    ifehl = 6
    
    Select Case index
        Case Is = 0     '** BAR **
            gbGutschOverBar = True
            SSCommand8(0).Enabled = False
            gcZahlMittel = "BA"
            giZahlArt = 8
            
            '*****Rest
            
            gcRueckgeld = Format$(dGeldAnKunden, "#####0.00")
            
            ifehl = 7
            If index < 3 Then '1
            
                TSSBerechnung
                SendeDaten2DruckerGutschein2WK20g 0, dZuZahlen, dGegeben, dWertRestGutschein
            End If
            
            ifehl = 8
            
            dGegeben = 0
            If index < 3 Then
                UpdateAFCStatGutscheinModul20 dZuZahlen, dEingereichteGutscheine, dNochOffen, dWertRestGutschein
                InsertAFCBuchGutscheinModul20 dEingereichteGutscheine
                If CheckofP = True Then
                    InsertProvision
                End If
                
                If CheckofX = True Then
                    InsertXMarkierung
                End If
                If dwertGutverkauf > 0 Then
                
                    If dEingereichteGutscheine < dwertGutverkauf Then
                        dwertGutverkauf = dEingereichteGutscheine
                    End If
                    updateafcstat "GUTSCHGUTSCH", dwertGutverkauf, gcKasNum
                    updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
                End If
                dwertGutverkauf = 0
                
            End If
            ifehl = 9

            
            If gbGutschUNDlastschrift = False Then
                ReInitDialog20WK20g
                If index = 0 Then
                    gcRueckgeld = Format$(dGeldAnKunden, "#####0.00")
                    frmWK20a.Show 1
                End If
            End If
            
            ifehl = 10
            
            If index <> 3 Then
                Unload frmWK20g
            Else
                frmWK20g.Visible = False
                frmWKL20.List11.SetFocus
            End If
            
            ifehl = 11
            'Ende Rest
        Case Is = 1     '** Scheck **
            gcZahlMittel = "SC"
            giZahlArt = 6
            
            '*****Rest
            If index < 3 Then '1
            
                TSSBerechnung
                SendeDaten2DruckerGutschein2WK20g 0, dZuZahlen, dGegeben, dWertRestGutschein
            End If
            
            ifehl = 12
            dGegeben = 0 'Variable auf 0 setzen WK678
            If index < 3 Then
                UpdateAFCStatGutscheinModul20 dZuZahlen, dEingereichteGutscheine, dNochOffen, dWertRestGutschein
                InsertAFCBuchGutscheinModul20 dEingereichteGutscheine
                If CheckofP = True Then
                    InsertProvision
                End If
                
                If CheckofX = True Then
                    InsertXMarkierung
                End If
                
                If dwertGutverkauf > 0 Then
                    updateafcstat "GUTSCHGUTSCH", dwertGutverkauf, gcKasNum
                    updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
                End If
                dwertGutverkauf = 0
            End If
            
            ifehl = 13
            If gbGutschUNDlastschrift = False Then
                ReInitDialog20WK20g
                If index = 0 Then
                    gcRueckgeld = Format$(dGeldAnKunden, "#####0.00")
                    frmWK20a.Show 1
                End If
            End If
            
            ifehl = 14
            If index <> 3 Then
                Unload frmWK20g
            Else
                frmWK20g.Visible = False
                frmWKL20.List11.SetFocus
            End If
            'Ende Rest
            ifehl = 15
            
            
            
            
        Case Is = 2     '** Karte **
            '************************************************************
            If gbEcash Then
            
                ifehl = 16
                Select Case gsEPartner
                    Case Is = "ADT"
                        If gsAdtVerfahren = "XML" Then
                        
                            Dim hwnd&
                            Dim Y As String
                            Dim result&
                            Dim Title$
                            
                            Y = "SECpos Pay" '  (Terminal-ID: " & gsTerminalid & ")"
                                                
                            hwnd = GetWindow(Me.hwnd, GW_HWNDFIRST)
                        
                            Do
                                result = GetWindowTextLength(hwnd) + 1
                                Title = Space(result)
                                result = GetWindowText(hwnd, Title, result)
                                Title = Left$(Title, Len(Title) - 1)
                        
                                If InStr(1, Title, Y) Then
                        '            MsgBox hwnd
                                    SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
                        
                                End If
                        
                                hwnd = GetWindow(hwnd, GW_HWNDNEXT)
                            Loop Until hwnd = 0
                        End If
                End Select
            End If
        
            ifehl = 17
            '********************************************************
            If dGegeben <> dNochOffen Then
                sNochOffen = Format$(dNochOffen, "#####0.00")
                MsgBox "Sie können nur " & sNochOffen & " EUR per Karte abbuchen.", vbInformation, "Winkiss Hinweis:"
                Text3.SetFocus
                Text3.Text = sNochOffen
                Screen.MousePointer = 0
                Exit Sub
            End If
            giZahlArt = 17
            
            ifehl = 18
            frmWKL29.Show 1      '** Rückgabewert steht in <gcKreditkarte> **
                
            If gbErfolg = False Then
                Screen.MousePointer = 0
                Exit Sub
            Else
                ifehl = 19
                AndersZahlungWKL20g (17)
                    
                If gbGutschUNDlastschrift = False Then
                    If index = 0 Then
                        gcRueckgeld = Format$(dGeldAnKunden, "#####0.00")
                        frmWK20a.Show 1
                    End If
                End If
                    
                ifehl = 20
                        
                If index = 2 Then
                
                Else
                    If index <> 3 Then
                        Unload frmWK20g
                    Else
                        frmWK20g.Visible = False
                        frmWKL20.List11.SetFocus
                    End If
                End If
                'Ende Rest
                ifehl = 21
            End If
            
        Case Is = 3  '** EC-Lastschrift **
            ifehl = 22
            If dGegeben <> dNochOffen Then
                sNochOffen = Format$(dNochOffen, "#####0.00")
                MsgBox "Sie können nur " & sNochOffen & " EUR per EC - Lastschrift abbuchen.", vbInformation, "Winkiss Hinweis:"
                Text3.SetFocus
                Text3.Text = sNochOffen
                Screen.MousePointer = 0
                Exit Sub
            End If
            ifehl = 23
            '*****
            If gbZwangsKdNr Then
                If frmWKL20.Label2(7).Caption = "0" Then
                    MsgBox "Verkauf ohne Kunden-Nr nicht möglich!", vbCritical, "STOP!"

                    Exit Sub
                End If
            End If
            
            ifehl = 24
            If frmWKL20.Label2(7).Caption <> "0" Then
                lRet = frmWKL20.fnPruefeKundenSperrungWKL20()
                If lRet <> 0 Then
                    MsgBox "Kunde ist für Zahlungen mit EC-Karte gesperrt!", vbCritical, "Winkiss Hinweis:"
                    Exit Sub
                End If
            End If
            ifehl = 25
            gcZahlMittel = "LS"
            giZahlArt = 1
            
            ctmp = Label5(2).Caption
            ctmp = fnMoveComma2Point$(ctmp)
            dWert = Val(ctmp)
            
            If gFirma.BLZ = "" Or gFirma.Konto = "" Or gFirma.BankName = "" Then
                MsgBox "Für Ihr Unternehmen ist keine Kontoverbindung angegeben!" & vbCrLf & "(->Service ->Einstellungen ->Unternehmensdaten)" & vbCrLf & "EC-Lastschrift nicht möglich!", vbCritical, "STOP!"
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            If dWert > 0 Then
                cZeile1 = "EC-Lastschrift:"
                cZeile2 = gcWaehrung & " " & Label5(2).Caption
                cZeile2 = Space$(20 - Len(cZeile2)) & cZeile2
                
                ifehl = 26
                ZeigeKundenDisplay cZeile1, cZeile2

                frmWKL20.Label11(0).Visible = True
                frmWKL20.Label11(4).Visible = False

                LeereDatenECKarteWKL20
'                frmWKL20.LeereDatenECKarteWKL20

                frmWKL20.List11.Clear
                frmWKL20.Label11(3).Caption = Label5(2).Caption & " " & gcWaehrung
                
                ifehl = 27
                frmWKL20.Frame18.Visible = True
                frmWKL20.Frame18.ZOrder 0
                gbGutschUNDlastschrift = True
                gbGutschein = True
                gdGutLastRest = dWert
                
                frmWKL20.MSComm1.CommPort = gVerbindung.iComPort
                frmWKL20.MSComm1.InputLen = 0
                frmWKL20.MSComm1.Settings = gVerbindung.cSettings
                frmWKL20.MSComm1.RThreshold = 1
                If Not frmWKL20.MSComm1.PortOpen = True Then
                    frmWKL20.MSComm1.PortOpen = True
                End If
                
                ifehl = 28

            ElseIf dWert < 0 Then
                MsgBox "EC-Lastschrift ist bei negativen Beträgen nicht möglich!", vbCritical, "STOP!"

            ElseIf dWert = 0 Then
                MsgBox "Der Endbetrag ist 0. EC-Lastschrift ist nicht möglich!", vbCritical, "STOP!"

            End If
            
            
            '*****Rest
            
            ifehl = 29
            If index < 3 Then '1
            
                TSSBerechnung
                SendeDaten2DruckerGutschein2WK20g 0, dZuZahlen, dGegeben, dWertRestGutschein
            End If
            
            ifehl = 30
            dGegeben = 0 'Variable auf 0 setzen WK678
            If index < 3 Then
                UpdateAFCStatGutscheinModul20 dZuZahlen, dEingereichteGutscheine, dNochOffen, dWertRestGutschein
                InsertAFCBuchGutscheinModul20 dEingereichteGutscheine
                If CheckofP = True Then
                    InsertProvision
                End If
                
                If CheckofX = True Then
                    InsertXMarkierung
                End If
                If dwertGutverkauf > 0 Then
                    updateafcstat "GUTSCHGUTSCH", dwertGutverkauf, gcKasNum
                    updateafcstat "ZHLGGUTSCH", (-1 * dwertGutverkauf), gcKasNum
                End If
                dwertGutverkauf = 0
            End If
            
            ifehl = 31
            If gbGutschUNDlastschrift = False Then
                ReInitDialog20WK20g
                If index = 0 Then
                    gcRueckgeld = Format$(dGeldAnKunden, "#####0.00")
                    frmWK20a.Show 1
                End If
            End If
            
            ifehl = 32
            
            If index <> 3 Then
                Unload frmWK20g
            Else
                frmWK20g.Visible = False
                frmWKL20.List11.SetFocus
            End If
            'Ende Rest
            
        Case Is = 4 '** Abbrechen **
        
            ifehl = 33
            Frame1.Visible = True
            Frame2.Visible = True
            Frame3.Visible = True
            Frame5.Visible = True
            Frame8.Visible = False
            dNochOffen = 0
            dZuZahlen = 0
            dGegeben = 0
            bFirstEingabe = False
            Screen.MousePointer = 0
            Exit Sub
    End Select
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Or err.Number = 8012 Then
        ifehl = 44
        Resume Next
    ElseIf err.Number = 400 Then
        ifehl = 45
        Resume Next
    ElseIf err.Number = 8002 Then
        ifehl = 46
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "SSCommand8_Click"
        Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. " & ifehl
        
        Fehlermeldung1
    
    End If
End Sub
Private Sub AndersZahlungWKL20g(iTaste As Integer)
    On Error GoTo LOKAL_ERROR
        
    Dim ctmp            As String
    Dim dSumme          As Double
    Dim dWert           As Double
    Dim bKommaVorhanden As Boolean
    
    
    
    
    giAndersZahlung = iTaste
    Label15.Caption = "Kartenverkauf"

    ctmp = Text3.Text
    bKommaVorhanden = False
    
    If InStr(ctmp, ",") Then
        bKommaVorhanden = True
    End If
    
    ctmp = fnMoveComma2Point$(ctmp)
    dSumme = Val(ctmp)
    
    If Not bKommaVorhanden Then
        dSumme = dSumme / 100
    End If

    Label6(5).Caption = Format$(dSumme, "#####0.00")
    Frame10.Visible = True
    Command5(1).Caption = "Zurück"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "AndersZahlungWKL20g"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
End Sub
Private Sub Text1_Change()
    On Error GoTo LOKAL_ERROR
    
'    checkgutschScan Text1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_Change"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyReturn Then
        SSCommand1_Click 1
        If bseekerfolg Then
            SSCommand2_Click 0
        Else

        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Text2_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = glSelBack1
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Text2_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text2.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text2_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1

End Sub
Private Sub Text3_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = glSelBack1
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Text3_LostFocus()
    On Error GoTo LOKAL_ERROR
    
    Text3.BackColor = vbWhite
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text3_LostFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
    
End Sub
Private Sub Text5_GotFocus()
    On Error GoTo LOKAL_ERROR
    
    Text5.BackColor = glSelBack1
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text5_GotFocus"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
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
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Bezahlen mit Gutschein auf. "
    
    Fehlermeldung1
End Sub


