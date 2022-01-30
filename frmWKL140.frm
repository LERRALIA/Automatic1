VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWKL140 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "MB festsetzen"
   ClientHeight    =   8625
   ClientLeft      =   1500
   ClientTop       =   2025
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   2  'Bildschirmmitte
   Begin sevCommand3.Command Command3 
      Height          =   375
      Index           =   11
      Left            =   11280
      TabIndex        =   36
      Top             =   120
      Width           =   375
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   5535
      Left            =   0
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   11895
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   6720
         MaxLength       =   3
         TabIndex        =   38
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   11
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Ausschließen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7011
         _Version        =   393216
         ForeColorSel    =   8454143
         FocusRect       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Index           =   1
         Left            =   7560
         TabIndex        =   40
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Übernehmen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Index           =   3
         Left            =   9480
         TabIndex        =   42
         Top             =   5160
         Visible         =   0   'False
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Löschen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MB Festsetzen:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   4200
         Width           =   6495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MB, der festgesetzt werden soll"
         Height          =   255
         Index           =   11
         Left            =   6720
         TabIndex        =   39
         Top             =   4320
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label0 
         BackColor       =   &H00800000&
         Caption         =   "Label0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   8160
         TabIndex        =   19
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   11775
      Begin VB.CheckBox Check2 
         Caption         =   "nur geräumte Artikel"
         Height          =   255
         Left            =   4440
         TabIndex        =   37
         Top             =   1320
         Width           =   2895
      End
      Begin sevCommand3.Command Command0 
         Height          =   255
         Index           =   6
         Left            =   8520
         TabIndex        =   35
         ToolTipText     =   "Kalender"
         Top             =   1320
         Width           =   1335
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
         Caption         =   "leeren"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   5760
         TabIndex        =   33
         Text            =   "123456"
         Top             =   120
         Width           =   1455
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   5
         Left            =   7320
         TabIndex        =   32
         ToolTipText     =   "Kalender"
         Top             =   120
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "123"
         Top             =   120
         Width           =   615
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   4
         Left            =   4680
         TabIndex        =   29
         ToolTipText     =   "Kalender"
         Top             =   120
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
      Begin VB.CheckBox Check1 
         Caption         =   "nur festgesetzte MB Artikel anzeigen"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   3855
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   7800
         MultiSelect     =   2  'Erweitert
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin sevCommand3.Command Command0 
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   26
         ToolTipText     =   "Kalender"
         Top             =   120
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
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   1
         Left            =   7200
         TabIndex        =   24
         ToolTipText     =   "Kalender"
         Top             =   600
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "123"
         Top             =   960
         Width           =   975
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   0
         Left            =   9480
         TabIndex        =   9
         ToolTipText     =   "Kalender"
         Top             =   600
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
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   2
         Left            =   6240
         TabIndex        =   8
         ToolTipText     =   "Kalender"
         Top             =   600
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
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   9000
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "123"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   7560
         MaxLength       =   13
         TabIndex        =   4
         Text            =   "1234567890123"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   5520
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "123456"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   120
         MaxLength       =   35
         TabIndex        =   0
         Text            =   "12345678901234567890123456789012345"
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3960
         MaxLength       =   13
         TabIndex        =   1
         Text            =   "1234567890123"
         Top             =   960
         Width           =   1575
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Index           =   0
         Left            =   9960
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Suchen"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Index           =   2
         Left            =   9960
         TabIndex        =   41
         Top             =   1200
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Übersicht"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Marke"
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
         Index           =   11
         Left            =   5160
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PGN"
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
         Index           =   10
         Left            =   3600
         TabIndex        =   31
         Top             =   240
         Width           =   495
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
         Left            =   1800
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   8
         Left            =   6720
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Artikelsuche"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   4
         Left            =   9120
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Lief.-Bestell-Nr."
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
         Index           =   3
         Left            =   7680
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Lief.-Nr"
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
         Index           =   2
         Left            =   5520
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Artikelbezeichnung"
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "ArtikelNr / EAN"
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
         Index           =   0
         Left            =   4080
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   2
      Left            =   9480
      TabIndex        =   7
      Top             =   8040
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Schließen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin sevCommand3.Command Command11 
      Height          =   360
      Left            =   10800
      TabIndex        =   43
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
      Picture         =   "frmWKL140.frx":0000
      PictureAlign    =   3
      Version3        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11640
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblAnzeige 
      Caption         =   "Anzeigetext"
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
      Left            =   120
      TabIndex        =   22
      Top             =   8040
      Width           =   9375
   End
   Begin VB.Label lblUeberschrift 
      BackStyle       =   0  'Transparent
      Caption         =   "MB festsetzen"
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
      TabIndex        =   20
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmWKL140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpaltennummerArtnr  As Byte
Private Function fnPruefeEingabeZENca() As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cSQL As String
    
    fnPruefeEingabeZENca = 1
    
    If Label4(32).Caption = "Farbauswahl" Then
        fnPruefeEingabeZENca = 0
        Exit Function
    End If
    
    For lcount = 0 To 7
        If Trim$(Text1(lcount).Text) <> "" Then
            fnPruefeEingabeZENca = 0
            Exit Function
        End If
    Next lcount
    
    If Check1.Value = vbChecked Then
        fnPruefeEingabeZENca = 0
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "fnPruefeEingabeZENca"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub FormatiereMSFlexgrid1ZENca()
    On Error GoTo LOKAL_ERROR
    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 29
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.FixedCols = 2
    
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 0
    MSFlexGrid1.ColWidth(0) = 700
    MSFlexGrid1.Text = "ArtNr."
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.ColWidth(1) = 3500
    MSFlexGrid1.Text = "Artikelbezeichnung"
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.ColWidth(2) = 1200
    MSFlexGrid1.Text = "G.Bestand"
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.ColWidth(3) = 900
    MSFlexGrid1.Text = "Kassen-VK"
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.ColWidth(4) = 900
    MSFlexGrid1.Text = "Listen-VK"
    
    MSFlexGrid1.Col = 5
    MSFlexGrid1.ColWidth(5) = 900
    MSFlexGrid1.Text = "Listen-EK"
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.ColWidth(6) = 900
    MSFlexGrid1.Text = "Schnitt-EK"
    
    MSFlexGrid1.Col = 7
    MSFlexGrid1.ColWidth(7) = 700
    MSFlexGrid1.Text = "LiefNr."
    
    MSFlexGrid1.Col = 8
    MSFlexGrid1.ColWidth(8) = 1300
    MSFlexGrid1.Text = "LiefBestNr"
    
    MSFlexGrid1.Col = 9
    MSFlexGrid1.ColWidth(9) = 1300
    MSFlexGrid1.Text = "EAN"
    
    MSFlexGrid1.Col = 10
    MSFlexGrid1.ColWidth(10) = 1300
    MSFlexGrid1.Text = "EAN-2"
    
    MSFlexGrid1.Col = 11
    MSFlexGrid1.ColWidth(11) = 1300
    MSFlexGrid1.Text = "EAN-3"
    
    MSFlexGrid1.Col = 12
    MSFlexGrid1.ColWidth(12) = 500
    MSFlexGrid1.Text = "Linie"
    
    MSFlexGrid1.Col = 13
    MSFlexGrid1.ColWidth(13) = 500
    MSFlexGrid1.Text = "RKZ"
    
    MSFlexGrid1.Col = 14
    MSFlexGrid1.ColWidth(14) = 600
    MSFlexGrid1.Text = "MWSt"
    
    MSFlexGrid1.Col = 15
    MSFlexGrid1.ColWidth(15) = 900
    MSFlexGrid1.Text = "MinBestell"
    
    MSFlexGrid1.Col = 16
    MSFlexGrid1.ColWidth(16) = 1000
    MSFlexGrid1.Text = "MinBestand"
    
    MSFlexGrid1.Col = 17
    MSFlexGrid1.ColWidth(17) = 800
    MSFlexGrid1.Text = "Inhalt"
    
    MSFlexGrid1.Col = 18
    MSFlexGrid1.ColWidth(18) = 800
    MSFlexGrid1.Text = "Einheit"
    
    MSFlexGrid1.Col = 19
    MSFlexGrid1.ColWidth(19) = 900
    MSFlexGrid1.Text = "Grundpreis"
    
    MSFlexGrid1.Col = 20
    MSFlexGrid1.ColWidth(20) = 800
    MSFlexGrid1.Text = "Rabatt"
    
    MSFlexGrid1.Col = 21
    MSFlexGrid1.ColWidth(21) = 800
    MSFlexGrid1.Text = "Geführt"
    
    MSFlexGrid1.Col = 22
    MSFlexGrid1.ColWidth(22) = 800
    MSFlexGrid1.Text = "Bonus"
    
    MSFlexGrid1.Col = 23
    MSFlexGrid1.ColWidth(23) = 800
    MSFlexGrid1.Text = "AGN"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FormatiereMSFlexgrid1ZENca"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub SucheArtikelZENca()
    On Error GoTo LOKAL_ERROR
    
    Dim iRet As Integer
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim lAnz As Long
    Dim lAkt As Long
    Dim ctmp        As String
    Dim dWert       As Double
    Dim j           As Integer
    Dim byteOrder   As Byte
    Dim brkz    As Boolean
    
    
    iRet = fnPruefeEingabeZENca()
    If iRet <> 0 Then
        anzeige "rot", "Bitte mindestens ein Suchkriterium angeben!", lblanzeige
        Text1(0).SetFocus
        Exit Sub
    End If
    
    anzeige "normal", "Artikel werden ermittelt, bitte warten...", lblanzeige
    
    Tabcheck "MBFEST"
    FormatGridOverTablay "MBFEST"

    With MSFlexGrid1
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
            aBreite(j) = TextWidth(.TextMatrix(0, j)) '* 1.8
        Next j
    End With
    
    MSFlexGrid1.Visible = False
    SSCommand2(0).Visible = False
    SSCommand2(1).Visible = False
    SSCommand2(3).Visible = False
    Label3(11).Visible = False
    Text5.Visible = False
    
    byteOrder = 1

    If Check2.Value = vbChecked Then
        brkz = True
    Else
        brkz = False
    End If
    
    If Check1.Value = vbChecked Then
    
        cSQL = fnBildeSQLZENartFMb(srechnertab, Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(5).Text, _
        Text1(3).Text, Text1(4).Text, byteOrder, Label4(32).Tag, List3, True, Text1(6).Text, Text1(7).Text, brkz)
    Else
        cSQL = fnBildeSQLZENartFMb(srechnertab, Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(5).Text, _
        Text1(3).Text, Text1(4).Text, byteOrder, Label4(32).Tag, List3, False, Text1(6).Text, Text1(7).Text, brkz)
    End If
    
    GridFuellen cSQL
    
    ermittlespalten
    
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    
    SSCommand2(0).Visible = True
    SSCommand2(1).Visible = True
    SSCommand2(3).Visible = True
    Label3(11).Visible = True
    Text5.Visible = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SucheArtikelZENca"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Tabellenbreiteanpassen(gridx As MSFlexGrid, siEigFak As Single)
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
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Bei Anpassen der Tabellenbreite ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub GridFuellen(cSQL As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim iRet        As Integer
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim counter     As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim lMax        As Long
    
    If cSQL = "" Then
        Exit Sub
    End If
    Set rsrs = gdApp.OpenRecordset(cSQL)
    
    With MSFlexGrid1
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lMax = rsrs.RecordCount
    
        anzeige "normal", "Es werden " & lMax & " Artikel angezeigt...", lblanzeige
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
                        Case Is = "L-EK", "K-VK", "L-VK.", "Schnitt-EK", "KVK neu", "Lagerumschlag"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = Format$(sWert, "####0.00")
                            
                        Case Is = "Preisschutz", "Geführt"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "N"
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                        Case Is = "Rabatt"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "J"
                            End If
                            .Row = lrow
                            .Text = sWert
                            
                        Case Is = "AWM"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = sWert
                            FaerbenFlex sWert, MSFlexGrid1, 0, CInt(lrow)
                         
                        Case Is = "MinBest"
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = "0"
                            End If
                            .Row = lrow
                            .Text = sWert
                        
                        Case Else
                            If Not IsNull(rsrs(sSpaltenbez(i))) Then
                                sWert = rsrs(sSpaltenbez(i))
                            Else
                                sWert = ""
                            End If
                            .Row = lrow
                            .Text = sWert
                    End Select
                    
                    If TextWidth(.TextMatrix(lrow, i)) > aBreite(i) Then
                        aBreite(i) = TextWidth(.TextMatrix(lrow, i))
                    End If
                    
                End If
            Next i
                                
            rsrs.MoveNext
        Loop
        
        Frame2.Visible = True
        anzeige "normal", "Es wurden " & lMax & " Artikel ermittelt.", lblanzeige
    Else
        Frame2.Visible = False
        anzeige "rot", "Es wurden keine Artikel ermittelt.", lblanzeige
    End If
    
    For i = 0 To byAnzahlSpalten - 1
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.8
    Next i
        
    rsrs.Close
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
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Private Sub Command0_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    Dim lcount As Long
    
    Select Case Index
    
        Case Is = 0
            Text1_KeyUp 4, vbKeyF2, 0
            
        Case Is = 2
            Text1_KeyUp 2, vbKeyF2, 0
            
        Case Is = 1
            Text1_KeyUp 5, vbKeyF2, 0
            
        Case Is = 4
            Text1_KeyUp 6, vbKeyF2, 0
            
        Case Is = 5
            Text1_KeyUp 7, vbKeyF2, 0
            
        Case Is = 3
    
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
        Case 6
            leer
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub



Private Sub Command11_Click()
On Error GoTo LOKAL_ERROR

    gsZSpalte = "Artnr"
    gstab = "MBFEST" 'Artbea"
    frmWKL36.Show 1
    'fertig
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command11_Click"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command3_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
    
    Case 11
        gsHelpstring = "MB festsetzen"
        frmWKL110.Show 1
        
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command3_Click"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Form_Load()
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    PositionierenZENca
    
    Modul6.Skalieren Me, True, True: Modul6.Schrift Me: Modul6.Log Me
    Modul6.Farbform Me, lblUeberschrift
    
    LogtoStart Me
    
    Dim slbl As String
    
    slbl = "MB Festsetzen:" & vbCrLf
    slbl = slbl & "Artikel zur Anzeige bringen. Alle Artikel, die in der Tabellenansicht enthalten sind, "
    slbl = slbl & "können jetzt mit einer Meldemenge durch Klick auf 'Übernehmen' versehen werden. "
    slbl = slbl & "Achtung: Sollten einige Artikel von dieser Aktion ausgeschlossen werden, so markieren Sie diese einzeln und klicken auf 'Ausschließen'."
    
    slbl = slbl & vbCrLf & vbCrLf
    
    slbl = slbl & "MB Festsetzung löschen:" & vbCrLf
    slbl = slbl & "Artikel zur Anzeige bringen. (Haken: 'nur festegesetze MB anzeigen' setzen) Mit Klick auf 'Löschen' können Sie  "
    slbl = slbl & "deren MB-Festsetzung aufheben. (betroffen sind alle Artikel in der Tabellenansicht)"
    slbl = slbl & "Achtung: Sollten einige Artikel von dieser Aktion ausgeschlossen werden, so markieren Sie diese einzeln und klicken auf 'Ausschließen'."
    
    
    Label3(0).Caption = slbl
    
    
    leer
    
    Frame2.Visible = True
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Form_Load"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase$(sSpaltenbez(i))
            Case Is = "ARTNR"
                SpaltennummerArtnr = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub leer()
    On Error GoTo LOKAL_ERROR
    
    anzeige "normal", "", lblanzeige
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    Text1(6).Text = ""
    Text1(7).Text = ""
    
    List3.Clear
    List3.Visible = False
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "leer"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Function ermlfnr(cART As String, sTab As String) As Long
    On Error GoTo LOKAL_ERROR
    
    ermlfnr = 0
    
    If cART = "" Then
        Exit Function
    End If
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select lfnr from " & sTab & " where ARTNR = " & cART & " "
    Set rsINB = gdApp.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!lfnr) Then
            ermlfnr = rsINB!lfnr
        End If
    End If
    rsINB.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermlfnr"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Sub PositionierenZENca()
    On Error GoTo LOKAL_ERROR
    
    With Frame1
        .Top = 720
        .Left = 0
        .Height = 1695
        .Width = 11895
        .BorderStyle = 0
    End With
    
    With Frame2
        .Top = 2280
        .Left = 0
        .Width = 11895
        .Height = 5655
        .BorderStyle = 0
    End With

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PositionierenZENca"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Label4_dblClick(Index As Integer)
On Error GoTo LOKAL_ERROR

If Index = 32 Then
    Label4(Index).Caption = "alle Farben"
    Label4(Index).Tag = ""
    Label4(Index).BackColor = Label1(1).BackColor
    Label4(Index).ForeColor = Label1(1).ForeColor
    
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Label4_dblClick"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR

Dim sSpaltenbez As String

If MSFlexGrid1.Row = 1 Then

    MSFlexGrid1.Row = 0
    sSpaltenbez = MSFlexGrid1.Text
    If sSpaltenbez <> "" Then
        sortierenArtt1 sSpaltenbez, "MBFEST", srechnertab
    End If
    sSpaltenbez = ""
    
    GridFuellen "Select * from " & srechnertab & " order by lfnr"
    Tabellenbreiteanpassen MSFlexGrid1, 1.25 * gdTabfak
    MSFlexGrid1.Row = 2
    
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub sortierenArtt1(sSBez As String, sTab As String, sglobatab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSortSpalte     As String
    Dim sSQL            As String
    Dim rs              As Recordset
    
    sSQL = " Select spaltenbez from tablay where Tabname = '" & sTab & "' "
    sSQL = sSQL & " and spaltenna = '" & sSBez & "' "
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!Spaltenbez) Then
            sSortSpalte = rs!Spaltenbez
        End If
    End If
    rs.Close
    
    'sortiere neu
    sSQL = "alter table " & sglobatab & " drop column LFNR"
    gdApp.Execute sSQL, dbFailOnError
    
    loeschNEW sglobatab & "2", gdApp
    
    sSQL = " Select * into " & sglobatab & "2 from  " & sglobatab
    
    Dim corder As String
    corder = " order by " & sSortSpalte
    
    If byteSortReihen = 2 Then
        byteSortReihen = 1
        corder = corder & " asc "
    ElseIf byteSortReihen = 1 Then
        byteSortReihen = 2
        corder = corder & " desc "
    End If
    sSQL = sSQL & corder
    gdApp.Execute sSQL, dbFailOnError
    
    loeschNEW sglobatab, gdApp

    sSQL = " Select * into " & sglobatab & " from  " & sglobatab & "2"
    sSQL = sSQL & corder
    gdApp.Execute sSQL, dbFailOnError
    
    SpalteAnfuegenNEW sglobatab, "LFNR", "Autoincrement", gdApp
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "sortierenArtt1"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
'Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo LOKAL_ERROR
'
'    Select Case KeyCode
'
'        Case Is = vbKeyReturn
'            SSCommand2_Click 0
'        Case Is = vbKeyEscape
'            SSCommand1_Click 2
'        Case Is = vbKeyF2
'            gsARTNR = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
'
'            If Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr), 1) = "X" Then
'                gsARTNR = Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr), 6)
'            End If
'
'            If gsARTNR <> "" Then
'                If IsNumeric(gsARTNR) Then
'                    gllfnr = 0
'                    frmZENcb.Show 1
'                    Me.Refresh
'                End If
'
'                Screen.MousePointer = 0
'            End If
'
'        Case Is = vbKeyF3
'            gsARTNRFiliale = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr)
'
'            If Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr), 1) = "X" Then
'                gsARTNRFiliale = Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, SpaltennummerArtnr), 6)
'            End If
'
'
'
'
'            If IsNumeric(gsARTNRFiliale) Then
'                If gbBestinZ Then
'                    frmZENcg.Show 1
'                Else
'                    frmZENcf.Show 1
'                End If
'            Else
'                gsARTNRFiliale = ""
'            End If
'
'
'    End Select
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = Me.name
'    Fehler.gsFunktion = "MSFlexGrid1_KeyUp"
'    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'End Sub



Private Sub SSCommand1_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "LPROTOK"
    
    Screen.MousePointer = 11
    
    Select Case Index
        Case Is = 0     'Suchen
            SucheArtikelZENca
            If MSFlexGrid1.Visible = True Then
                MSFlexGrid1.SetFocus
                MSFlexGrid1.Row = 2
            End If
        Case Is = 2     'Schließen
            loeschNEW srechnertab, gdBase
            loeschNEW srechnertab, gdApp
            Unload frmWKL140
    End Select
    
    Screen.MousePointer = vbDefault
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand1_Click"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub SSCommand2_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lcol        As Long
    
    Select Case Index
    
        Case Is = 0     'Entfernen
            If MSFlexGrid1.RowSel > 1 Then
                FlexGrid_Update MSFlexGrid1
            End If
            lrow = MSFlexGrid1.Row
            lcol = MSFlexGrid1.Col

            MSFlexGrid1.Col = lcol
            MSFlexGrid1.Row = lrow
            MSFlexGrid1.SetFocus
        Case 1 'Übernahme
            DelEntfernt
            Uebernahme
        Case 2 'Übersicht
            Uebersicht
        Case 3 'Löschen mit Filialauswahl
            If Not NewTableSuchenDBKombi("MBORDERDEL", gdBase) Then
                CreateTableT2 "MBORDERDEL", gdBase
            End If
            DelEntfernt
            Loeschen
    End Select
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_Click"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Private Sub Loeschen()
On Error GoTo LOKAL_ERROR

    Dim iMB             As Integer
    Dim sSQL            As String
    Dim cPfad           As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad 'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    loeschNEW srechnertab, gdBase
    TransferTab gdApp, cPfad & "kissdata.mdb", srechnertab

    anzeige "normal", "", lblanzeige
    
    sSQL = " Delete from MBORDER where artnr in ( "
    sSQL = sSQL & " select ARTNR"
    sSQL = sSQL & " from " & srechnertab
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into MBORDERDEL"
    sSQL = sSQL & " select ARTNR"
    sSQL = sSQL & " , " & gcFilNr & " as FILIALE "
    sSQL = sSQL & " , '" & DateValue(Now) & "' as Lastdate"
    sSQL = sSQL & " , " & gcBedienerNr & " as AENDWER "
    sSQL = sSQL & " , false as sendok"
    sSQL = sSQL & " from " & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "", lblanzeige

    Screen.MousePointer = 0
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Loeschen"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub DelEntfernt()
On Error GoTo LOKAL_ERROR

    Dim lAnzRows    As Long
    Dim lcount      As Long
    Dim sSQL        As String
    Dim cART        As String
    
    Screen.MousePointer = 11
    
    lAnzRows = MSFlexGrid1.Rows
    lAnzRows = lAnzRows - 1

    For lcount = 2 To lAnzRows
       
        cART = MSFlexGrid1.TextMatrix(lcount, SpaltennummerArtnr)
        
        If Left(MSFlexGrid1.TextMatrix(lcount, SpaltennummerArtnr), 1) = "X" Then
            cART = Right(MSFlexGrid1.TextMatrix(lcount, SpaltennummerArtnr), 6)
            sSQL = " Delete from " & srechnertab & " where artnr = " & cART
            gdApp.Execute sSQL, dbFailOnError
        
        End If
    
    Next lcount
    
   

    Screen.MousePointer = 0
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "DelEntfernt"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Uebernahme()
On Error GoTo LOKAL_ERROR

    Dim i       As Integer
    Dim iMB     As Integer
    Dim sSQL    As String
    Dim cPfad   As String
    
    Screen.MousePointer = 11
    
    
    If Text5.Text = "" Then
        anzeige "rot", "Bitte einen Wert eingeben!", lblanzeige
        Text5.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(Text5.Text) = False Then
        anzeige "rot", "Bitte einen Zahlenwert eingeben!", lblanzeige
        Text5.SetFocus
        Exit Sub
    End If
    
    
    
    cPfad = gcDBPfad 'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    loeschNEW srechnertab, gdBase
    TransferTab gdApp, cPfad & "kissdata.mdb", srechnertab

    anzeige "normal", "Das Festsetzen der Mindestbestände beginnt...", lblanzeige
    
    
    iMB = CInt(Text5.Text)
    
    If iMB < 0 Then iMB = 0

    sSQL = " Delete from MBORDER where artnr in ( "
    sSQL = sSQL & " select ARTNR"
    sSQL = sSQL & " from " & srechnertab
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
        
    sSQL = " Insert into MBORDER"
    sSQL = sSQL & " select ARTNR"
    sSQL = sSQL & " , " & iMB & " as MB "
    sSQL = sSQL & " , '" & DateValue(Now) & "' as Lastdate"
    sSQL = sSQL & " , " & gcBedienerNr & " as AENDWER "
    sSQL = sSQL & " , false as sendok"
    sSQL = sSQL & " from " & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ARTIKEL inner join MBORDER on ARTIKEL.artnr = MBORDER.artnr "
    sSQL = sSQL & " set ARTIKEL.minbest = MBORDER.MB "
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Das Festsetzen der Mindestbestände ist beendet.", lblanzeige

    Screen.MousePointer = 0
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebernahme"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Uebersicht()
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    Screen.MousePointer = 11
    
    anzeige "normal", "", lblanzeige
    
    loeschNEW "MBORDERPR", gdBase
    CreateTableT2 "MBORDERPR", gdBase
    
    sSQL = " Insert into MBORDERPR"
    sSQL = sSQL & " select * "
    sSQL = sSQL & " from MBORDER "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update MBORDERPR inner join artikel on MBORDERPR.Artnr = artikel.Artnr "
    sSQL = sSQL & " set MBORDERPR.bezeich = artikel.bezeich "
    sSQL = sSQL & " , MBORDERPR.kvkpr1 = artikel.kvkpr1 "
    sSQL = sSQL & " , MBORDERPR.libesnr = artikel.libesnr "
    sSQL = sSQL & " , MBORDERPR.linr = artikel.linr "
    sSQL = sSQL & " , MBORDERPR.RKZ = artikel.RKZ "
    sSQL = sSQL & " , MBORDERPR.AUFDAT = artikel.AUFDAT "
    sSQL = sSQL & " , MBORDERPR.farbnr = val(artikel.awm) "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("MBORDERPR")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

        If Not IsNull(rsrs!artnr) Then
            rsrs.Edit
            rsrs!ERSTDAT = ErmFirstZugang(rsrs!artnr)
            rsrs.Update
        End If

        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
'    Anzeige "normal", "9. Schritt...", lblAnzeige
    
    sSQL = "Update MBORDERPR set NEU = 'N' where erstdat > datevalue(now) - 90 "
    gdBase.Execute sSQL, dbFailOnError
    
    BringFarbeInsSpiel "MBORDERPR", gdBase
    
    sSQL = "Update MBORDERPR inner join Lisrt on MBORDERPR.linr = lisrt.linr "
    sSQL = sSQL & " set MBORDERPR.liefbez = Lisrt.liefbez"
    gdBase.Execute sSQL, dbFailOnError
    
    anzeige "normal", "Druckvorschau wird erstellt...", lblanzeige
    reportbildschirm "", "aWkl140"

    Screen.MousePointer = 0
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Uebersicht"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub FlexGrid_Update(oGrid As MSFlexGrid)
On Error GoTo LOKAL_ERROR

    Dim nRow As Long
    Dim nCol As Long
    Dim nRowSel As Long
    Dim nColSel As Long
    Dim nDelRow As Long
    Dim lBig As Long

    With oGrid
        ' aktuelle Selektion merken
        nRow = .Row
        nCol = .Col
        nRowSel = .RowSel
        nColSel = .ColSel
        
        If nRow > nRowSel Then
            lBig = nRow
            nDelRow = nRowSel - 1
        Else
            lBig = nRowSel
            nDelRow = nRow - 1
        End If

        Do While nDelRow < lBig
            nDelRow = nDelRow + 1
            If nDelRow > 1 Then
                If Left(.TextMatrix(nDelRow, SpaltennummerArtnr), 1) <> "X" Then
                    .TextMatrix(nDelRow, SpaltennummerArtnr) = "X " & .TextMatrix(nDelRow, SpaltennummerArtnr)
                End If
            End If
        Loop
    End With
  
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "FlexGrid_Update"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub SSCommand2_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR
    
    If KeyCode = vbKeyEscape Then
        SSCommand1_Click 2
    ElseIf KeyCode = vbKeyA Then
        SSCommand2_Click 0
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SSCommand2_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    On Error GoTo LOKAL_ERROR

    Text1(Index).BackColor = glSelBack1
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
   
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
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
        Case 0, 2, 4, 5   'ARTNR, EAN, LIEFNR, ARTGRU,linie
            cValid = "1234567890" & Chr$(8)
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
        Case 1, 3       'BEZEICH, LIBESNR
            cValid = gcUPPER & gcLower & gcNUM & Chr$(8) & Chr$(32) & Chr(42) 'Leer *
            cValid = cValid & Chr(38) & Chr(45) & Chr(46) '& - .
            cValid = cValid & "+äÄÜüÖöß#"
            
            If InStr(cValid, cZeichen) = 0 Then
                KeyAscii = 0
            End If
            'alle Zeichen erlaubt
        
            
    End Select
        
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sAuswahlfeld As String
    Dim ctmp As String
    Dim lcount As Long
    
    If KeyCode = vbKeyReturn Then
        SSCommand1_Click 0
    End If
    
    If KeyCode = vbKeyEscape Then
        SSCommand1_Click 2
    End If

    
    If KeyCode = vbKeyF2 Then
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        Select Case Index
            Case Is = 2
                gF2Prompt.cFeld = "LINR"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text1(Index).Text = gF2Prompt.cWahl
                    End If
                End If
            
            Case Is = 4
                gF2Prompt.cFeld = "AGN"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text1(Index).Text = gF2Prompt.cWahl
                    End If
                End If
            Case Is = 6
                gF2Prompt.cFeld = "PGN"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text1(Index).Text = gF2Prompt.cWahl
                    End If
                End If
            Case Is = 7
                gF2Prompt.cFeld = "MARKE"
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    If gF2Prompt.cWahl <> "" Then
                        Text1(Index).Text = gF2Prompt.cWahl
                    End If
                End If
            Case 5
            
                ctmp = Text1(2).Text
                ctmp = Trim$(ctmp)
                If ctmp = "" Then
                    ctmp = Text1(7).Text
                    ctmp = Trim$(ctmp)
                    If ctmp = "" Then
                        anzeige "Rot", "Bitte einen Lieferanten oder eine Marke angeben!", lblanzeige
                        Text1(2).SetFocus
                        Exit Sub
                    Else
                        sAuswahlfeld = "MARKE"
                    End If
                Else
                    sAuswahlfeld = "LINR"
                End If
                
                gF2Prompt.bMultiple = True
                gF2Prompt.cFeld = "LPZ"
                gF2Prompt.cWert = ctmp
                gF2Prompt.cEsFeld = sAuswahlfeld
                
                If gF2Prompt.cFeld <> "" Then
                    frmWK00a.Show 1
                    
                    List3.Visible = False
                    List3.Clear
                    For lcount = 0 To 100
                        If lcount > 0 And gF2Prompt.cArray(lcount) <> "" Then
                            List3.Visible = True
                            Text1(Index).Text = ""
                            
                            If gF2Prompt.cArray(lcount) <> "" Then
                                List3.AddItem gF2Prompt.cArray(lcount) & Space(50) & Right(gF2Prompt.cArray(lcount), 6)
                            End If
                        Else
                            If gF2Prompt.cArray(lcount) <> "" Then
                                List3.AddItem gF2Prompt.cArray(lcount)
                                Text1(Index).Text = Left$(gF2Prompt.cArray(lcount), 3)
                            End If
                        End If
                    Next lcount
                End If
        End Select
        Text1(Index).SetFocus
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Text1_KeyUp"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
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
    Fehler.gsFunktion = "Text5_KeyPress"
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
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
    Fehler.gsFehlertext = "Im Programmteil MB festsetzen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
