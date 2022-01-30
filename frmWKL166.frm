VERSION 5.00
Object = "{7D622DE6-0ABC-471E-9234-97DEC5E0A708}#3.8#0"; "sevCmd3.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWKL166 
   Caption         =   "Excel Import"
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
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'Kein
      Caption         =   "Protokolle"
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
      Height          =   8295
      Left            =   -1200
      TabIndex        =   117
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   2
         Left            =   9480
         TabIndex        =   122
         Top             =   3960
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
         Caption         =   "Preisänderungen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   121
         Top             =   3360
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
         Caption         =   "neue Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command4 
         Height          =   495
         Index           =   0
         Left            =   9480
         TabIndex        =   120
         Top             =   2760
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
         Caption         =   "alle Artikel"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox CG 
         Caption         =   "nur geführte Artikel"
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
         Left            =   9480
         TabIndex        =   119
         Top             =   2280
         Value           =   1  'Aktiviert
         Width           =   2055
      End
      Begin sevCommand3.Command Command1 
         Height          =   495
         Index           =   7
         Left            =   9480
         TabIndex        =   118
         Top             =   5640
         Width           =   2040
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
      Begin VB.Label Label5 
         Caption         =   $"frmWKL166.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   125
         Top             =   600
         Width           =   9015
      End
      Begin VB.Label Label2 
         Caption         =   "Protokolle"
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
         Index           =   9
         Left            =   0
         TabIndex        =   124
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Möchten sie nur Ihre geführten Artikel berücksichtigt haben, dann setzen sie hier den Haken."
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
         Index           =   1
         Left            =   0
         TabIndex        =   123
         Top             =   2280
         Width           =   9015
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   495
      Index           =   9
      Left            =   9600
      TabIndex        =   116
      Top             =   7200
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
      Caption         =   "Protokolle"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.PictureBox picprogress 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   9315
      TabIndex        =   102
      Top             =   7800
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C000&
      Caption         =   "Pricat"
      Height          =   7335
      Left            =   1080
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   10815
      Begin sevCommand3.Command Command1 
         Height          =   360
         Index           =   4
         Left            =   9600
         TabIndex        =   129
         Top             =   5160
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
         Picture         =   "frmWKL166.frx":0091
         PictureAlign    =   3
         Version3        =   -1  'True
      End
      Begin VB.TextBox Text2 
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
         Index           =   3
         Left            =   8520
         MaxLength       =   6
         TabIndex        =   127
         Top             =   5640
         Width           =   495
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   1
         Left            =   9120
         TabIndex        =   126
         ToolTipText     =   "Abschlag auf den Listeneinkaufspeis durchführen"
         Top             =   5640
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "K"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   0
         Left            =   9600
         TabIndex        =   87
         ToolTipText     =   "Runden"
         Top             =   5640
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "R"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   2
         Left            =   10080
         TabIndex        =   86
         Top             =   5640
         Width           =   1440
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command1 
         Height          =   375
         Index           =   5
         Left            =   10080
         TabIndex        =   84
         Top             =   5160
         Width           =   1440
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Caption         =   "Übernehmen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "AGN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   8
         Left            =   2880
         TabIndex        =   83
         Top             =   5040
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Linien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   82
         Top             =   5280
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Einkaufspreis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   81
         Top             =   5520
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Listenverkaufspreis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   80
         Top             =   5280
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "alle Neuheiten auf  ""GEFÜHRT"" "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   79
         Top             =   5520
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Mindestbestellmenge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   78
         Top             =   5280
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Mindestmenge(VPE)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   77
         Top             =   5040
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Notizen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   76
         Top             =   5520
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kassenverkaufspreis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   75
         Top             =   5040
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "eventuelle Zweitlieferanten löschen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   10
         Left            =   5400
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "alte Artikel räumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   11
         Left            =   9000
         TabIndex        =   73
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C000&
         Caption         =   "Bezeichnung übernehmen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Index           =   9
         Left            =   6840
         TabIndex        =   72
         Top             =   5040
         Value           =   1  'Aktiviert
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid PRIFlex 
         Height          =   4575
         Left            =   120
         TabIndex        =   85
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   8070
         _Version        =   393216
         ForeColorSel    =   8454143
         FocusRect       =   0
         HighLight       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "LEK Abschlag in %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   8520
         TabIndex        =   128
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Index           =   35
         Left            =   7680
         TabIndex        =   89
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Left            =   7680
         TabIndex        =   88
         Top             =   5400
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Zuordnung der Spalten"
      Height          =   6135
      Left            =   840
      TabIndex        =   47
      Top             =   7560
      Visible         =   0   'False
      Width           =   11655
      Begin sevCommand3.Command Command5 
         Height          =   225
         Index           =   2
         Left            =   7080
         TabIndex        =   103
         Top             =   1320
         Width           =   225
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
         Caption         =   "x"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   3840
         TabIndex        =   65
         Top             =   2400
         Width           =   3495
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
         Height          =   3690
         ItemData        =   "frmWKL166.frx":0723
         Left            =   120
         List            =   "frmWKL166.frx":0725
         TabIndex        =   56
         Top             =   840
         Width           =   3615
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   7
         Left            =   7320
         TabIndex        =   55
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
         Caption         =   "weiter"
         Enabled         =   0   'False
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   8
         Left            =   9480
         TabIndex        =   54
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         Left            =   7440
         TabIndex        =   53
         Top             =   840
         Width           =   4095
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   11
         Left            =   3840
         TabIndex        =   52
         Top             =   1680
         Width           =   3495
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty MenuFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
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
         Caption         =   " <-- Zuordnen -->"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   120
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   50
         Top             =   5040
         Width           =   3615
      End
      Begin VB.ListBox List12 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   5040
         TabIndex        =   49
         Top             =   3840
         Width           =   6495
      End
      Begin sevCommand3.Command Command5 
         Height          =   255
         Index           =   16
         Left            =   5040
         TabIndex        =   48
         Top             =   5400
         Width           =   1095
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
         Caption         =   "Löschen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.ListBox List7 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   4920
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Zusammenstellung"
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
         Index           =   2
         Left            =   3840
         TabIndex        =   92
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Zusammenstellung"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   91
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Zusammenstellung"
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
         Index           =   0
         Left            =   3840
         TabIndex        =   90
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Vorauswahlsspalte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   28
         Left            =   3840
         TabIndex        =   66
         Top             =   2640
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Vorauswahlspalte"
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
         Index           =   34
         Left            =   3840
         TabIndex        =   64
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "enthaltene Spalten im Tabellenblatt:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   7080
         TabIndex        =   63
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "ExcelTabelleblattname"
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
         Index           =   8
         Left            =   7080
         TabIndex        =   62
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Zuordnung der Spalten"
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
         Index           =   20
         Left            =   120
         TabIndex        =   61
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Anzahl Spalten"
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
         Index           =   15
         Left            =   120
         TabIndex        =   60
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Exceldateiname"
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
         Left            =   3960
         TabIndex        =   59
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "ausgewählte Exceldatei:"
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
         Index           =   12
         Left            =   3960
         TabIndex        =   58
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Zusammenstellung"
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
         Index           =   7
         Left            =   5040
         TabIndex        =   57
         Top             =   3600
         Width           =   2415
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFF80&
      Caption         =   "3. Zusammenstellung"
      Height          =   3495
      Left            =   10200
      TabIndex        =   36
      Top             =   5760
      Visible         =   0   'False
      Width           =   3375
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   15
         Left            =   1800
         TabIndex        =   46
         Top             =   5520
         Width           =   5415
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
         Caption         =   "Zusammenstellung neu erstellen"
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   20
         Left            =   7320
         TabIndex        =   39
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
         Caption         =   "weiter"
         Enabled         =   0   'False
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   19
         Left            =   9480
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List13 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   3120
         TabIndex        =   37
         Top             =   1200
         Width           =   8415
      End
      Begin VB.Label Label2 
         Caption         =   "enthaltene Spalten im Tabellenblatt:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   7080
         TabIndex        =   45
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "ExcelTabelleblattname"
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
         Index           =   30
         Left            =   7080
         TabIndex        =   44
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "3. Zusammenstellung"
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
         Index           =   29
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Exceldateiname"
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
         Index           =   27
         Left            =   3960
         TabIndex        =   42
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "ausgewählte Exceldatei:"
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
         Index           =   25
         Left            =   3960
         TabIndex        =   41
         Top             =   120
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Zusammenstellung"
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
         Left            =   3120
         TabIndex        =   40
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Welche Artikel möchten Sie übernehmen?"
      Height          =   3135
      Left            =   5760
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CheckBox Check6 
         Caption         =   "AGN Zuordnung"
         Height          =   255
         Left            =   8760
         TabIndex        =   115
         Top             =   3000
         Value           =   1  'Aktiviert
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   67
         Top             =   960
         Visible         =   0   'False
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   18
         FixedCols       =   2
         ForeColorSel    =   8454143
         FocusRect       =   0
         HighLight       =   2
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
      Begin VB.CheckBox Check5 
         Caption         =   "MwSt Einstellungen"
         Height          =   255
         Left            =   9000
         TabIndex        =   114
         Top             =   720
         Width           =   2175
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'Kein
         Height          =   855
         Left            =   9000
         TabIndex        =   107
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   109
            Text            =   "1"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   108
            Text            =   "2"
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "E"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   113
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "V"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   112
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "="
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   111
            Top             =   480
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "="
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   110
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "LK - Auflösung"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   720
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Lieferantenkürzel vor die Artikelbezeichnung stellen"
         Height          =   495
         Left            =   6720
         TabIndex        =   105
         Top             =   4440
         Value           =   1  'Aktiviert
         Width           =   3135
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
         Height          =   345
         Index           =   0
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   104
         Top             =   4920
         Width           =   1095
      End
      Begin sevCommand3.Command Command0 
         Height          =   345
         Index           =   4
         Left            =   7920
         TabIndex        =   96
         Top             =   3240
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
         Index           =   5
         Left            =   7920
         TabIndex        =   95
         Top             =   2040
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
         Height          =   345
         Index           =   4
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   94
         Top             =   3240
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
         Height          =   345
         Index           =   2
         Left            =   6720
         MaxLength       =   6
         TabIndex        =   93
         Top             =   2040
         Width           =   1095
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   21
         Left            =   6720
         TabIndex        =   70
         Top             =   840
         Width           =   1935
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
         Caption         =   "alle entfernen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   375
         Index           =   22
         Left            =   6720
         TabIndex        =   69
         Top             =   1320
         Width           =   1935
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
         Caption         =   "alle auswählen"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   18
         Left            =   7320
         TabIndex        =   34
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
         Caption         =   "weiter"
         Enabled         =   0   'False
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   17
         Left            =   9480
         TabIndex        =   33
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
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
         Height          =   615
         Index           =   5
         Left            =   6720
         TabIndex        =   101
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
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
         ForeColor       =   &H00C0C0C0&
         Height          =   495
         Index           =   3
         Left            =   6720
         TabIndex        =   100
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   1
         Left            =   6720
         TabIndex        =   99
         Top             =   1800
         Width           =   1095
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
         Index           =   0
         Left            =   6720
         TabIndex        =   98
         Top             =   3000
         Width           =   1095
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
         Left            =   5160
         TabIndex        =   97
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   68
         Top             =   5640
         Width           =   6495
      End
      Begin VB.Label Label2 
         Caption         =   "Welche Artikel möchten Sie übernehmen?"
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
         Index           =   26
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   7335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "4. Werte dieser Spalte"
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
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
         TabIndex        =   24
         Top             =   2280
         Width           =   7455
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   10
         Left            =   9480
         TabIndex        =   28
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.CheckBox Check8 
         Caption         =   "eindeutige anzeigen"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   7455
      End
      Begin VB.Label Label2 
         Caption         =   "enthaltene Werte in der  Spalte:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   6375
      End
      Begin VB.Label Label2 
         Caption         =   "Excelwerte"
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
         Index           =   21
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "4. Werte dieser Spalte"
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
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Anzahl Werte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   5760
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "2. Auswahl des Tabellenblattes"
      Height          =   3255
      Left            =   -240
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   3735
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   5
         Left            =   9480
         TabIndex        =   21
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
         Caption         =   "weiter"
         Enabled         =   0   'False
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   4
         Left            =   9480
         TabIndex        =   20
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
         Caption         =   "Zurück"
         PictureAlign    =   2
         PictureVisible  =   0   'False
         Version3        =   -1  'True
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Klicken Sie ein Tabellenblatt an und drücken dann 'weiter'!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   19
         Left            =   4560
         TabIndex        =   19
         Top             =   2040
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "enthaltene Tabellenblätter in der Datei:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   6855
      End
      Begin VB.Label Label2 
         Caption         =   "Exceldateiname"
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
         Index           =   17
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Anzahl Tabellenblätter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   16
         Top             =   4080
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "2. Auswahl des Tabellenblattes"
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "1. Auswahl der Exceldatei"
      Height          =   4455
      Left            =   600
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
      Begin sevCommand3.Command Command5 
         Height          =   495
         Index           =   6
         Left            =   9480
         TabIndex        =   22
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
         Caption         =   "weiter"
         Enabled         =   0   'False
         PictureAlign    =   2
         Version3        =   -1  'True
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   120
         Pattern         =   "*.pri"
         TabIndex        =   9
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Klicken Sie eine Datei an und drücken dann weiter!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   16
         Left            =   4920
         TabIndex        =   12
         Top             =   1320
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "1. Auswahl der Exceldatei"
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label2 
         Caption         =   "Anzahl Dateien"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   10
         Top             =   3600
         Width           =   3615
      End
   End
   Begin sevCommand3.Command Command5 
      Height          =   375
      Index           =   3
      Left            =   9240
      TabIndex        =   5
      Top             =   120
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
      Caption         =   "Pfad ändern"
      PictureAlign    =   2
      PictureVisible  =   0   'False
      Version3        =   -1  'True
   End
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin sevCommand3.Command Command5 
      Height          =   345
      Index           =   1
      Left            =   11400
      TabIndex        =   3
      Top             =   360
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
   Begin VB.Label Label2 
      Caption         =   "Pfad zu den Stammdaten im Excelformat:"
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
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Anzeige"
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
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   9255
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
      Caption         =   "Excel Import"
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
      Width           =   6255
   End
End
Attribute VB_Name = "frmWKL166"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sExcelpfad As String
Dim glartv As Long
Dim glartb As Long
Dim SpaltennummerArtnr  As Byte
Dim SpaltennummerBEZEICH  As Byte
Dim SpaltennummerLINR  As Byte
Dim SpaltennummerLPZ  As Byte
Dim SpaltennummerLIBESNR  As Byte
Dim SpaltennummerLEKPR  As Byte
Dim SpaltennummerVKPR  As Byte
Dim SpaltennummerKVKPR1  As Byte
Dim SpaltennummerMINBEST  As Byte
Dim SpaltennummerGEFUEHRT  As Byte
Dim SpaltennummerRABATT_OK  As Byte
Dim SpaltennummerPREISSCHU  As Byte
Dim SpaltennummerNOTIZEN  As Byte
Dim SpaltennummerAGN  As Byte
Dim SpaltennummerPGN  As Byte
Dim SpaltennummerRKZ  As Byte
Dim SpaltennummerEAN  As Byte
Dim SpaltennummerMINMEN  As Byte
Dim SpaltennummerMENGE  As Byte
Dim SpaltennummerBESTAND  As Byte
Dim SpaltennummerMWST  As Byte
Dim SpaltennummerMNOTIZEN  As Byte
Dim SpaltennummerKVKNEU As Byte
Dim SpaltennummerGROESSE As Byte
Private Sub Check2_Click()
On Error GoTo LOKAL_ERROR

    Speichervorauswahl Check2, Label2(5).Caption & " " & Label2(8).Caption, Label2(28)
    Check2.BackColor = Label2(5).BackColor
    
    If checkZuordnung(Label2(5).Caption & " " & Label2(8).Caption) Then
        'weiter aktivieren
        Command5(7).Enabled = True
        Command5(7).BackColor = vbRed
    Else
        Command5(7).Enabled = False
        Command5(7).BackColor = Command5(8).BackColor
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check2_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check3_Click()
On Error GoTo LOKAL_ERROR

    If Check3.Value = vbChecked Then
        Text1(0).Text = UCase(Trim(Left(ermLiefBez(CLng(Text1(2).Text)), 3)))
    Else
        Text1(0).Text = ""
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check3_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check4_Click()
On Error GoTo LOKAL_ERROR

    If Check4.Value = vbChecked Then
        BIOPURLK
        
    Else
        
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check4_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BIOPURLK()
    On Error GoTo LOKAL_ERROR
    
    Dim cVorauswahlspalte        As String
    
    If Not NewTableSuchenDBKombi("BIOPURLK", gdBase) Then 'das erste Mal
        CreateTableT2 "BIOPURLK", gdBase
        
        BIOPURLKFuellen
    End If
    
    cVorauswahlspalte = ermVorauswahl(Label2(5).Caption & " " & Label2(8).Caption)
    zeigVorauswahlWerte Label2(5), Label2(8), cVorauswahlspalte, True
    

    Me.Refresh
   
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BIOPURLK"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub BIOPURLKFuellen()
    On Error GoTo LOKAL_ERROR
    
    NachtragBIOLK "Aries", "ARI"
    NachtragBIOLK "Philips GmbH/Avent", "AVE"
    NachtragBIOLK "Alnavit", "AVT"
    NachtragBIOLK "Bauck Hof", "BAK"
    NachtragBIOLK "Kornkreis Bioland", "BBW"
    NachtragBIOLK "Primavita", "BFS"
    NachtragBIOLK "Bluegreen", "BGR"
    NachtragBIOLK "Barnhouse", "BHO"
    NachtragBIOLK "Bingenheimer", "BIN"
    NachtragBIOLK "Probijo", "Bit"
    NachtragBIOLK "Biovital Biostro", "BIV"
    NachtragBIOLK "Blauer Planet", "BPT"
    NachtragBIOLK "Biosun", "BSN"
    NachtragBIOLK "Beltane", "BTA"
    NachtragBIOLK "Bioturm", "BTM"
    NachtragBIOLK "alle Verlage", "BÜC"
    NachtragBIOLK "Byodo", "BYO"
    NachtragBIOLK "Charmy", "CHA"
    NachtragBIOLK "Clostermann", "CLO"
    NachtragBIOLK "C.m.D", "CMD"
    NachtragBIOLK "Cosmoveda", "Cos"
    NachtragBIOLK "Davert Mühle", "DAV"
    NachtragBIOLK "Beerenbauern", "dbb"
    NachtragBIOLK "Demeter Ferlderz", "DFE"
    NachtragBIOLK "Eco Plus", "ECP"
    NachtragBIOLK "Arche", "ECV"
    NachtragBIOLK "Nürnberger/ Eisbl", "EIS"
    NachtragBIOLK "Equimol", "EQU"
    NachtragBIOLK "Erdmannhauser", "ERD"
    NachtragBIOLK "Evers", "EVS"
    NachtragBIOLK "Faan", "FAA"
    NachtragBIOLK "Farfalla", "FAR"
    NachtragBIOLK "Fitne", "FIT"
    NachtragBIOLK "Flores Farm", "FLF"
    NachtragBIOLK "Rabenhorst/Flemming", "FLN"
    NachtragBIOLK "Fontaine", "FON"
    NachtragBIOLK "Frisetta", "FRI"
    NachtragBIOLK "m.Gebhardt", "GEB"
    NachtragBIOLK "Glafey", "GLG"
    NachtragBIOLK "Kit BV/ Golden Temp", "GOL"
    NachtragBIOLK "Govinda", "GOV"
    NachtragBIOLK "Dr.Grandel", "GRA"
    NachtragBIOLK "GSE Vertrieb", "GSE"
    NachtragBIOLK "Wala/Dr. Hauschka", "HAU"
    NachtragBIOLK "Heidelberger", "HEI"
    NachtragBIOLK "Herbaria", "HER"
    NachtragBIOLK "Heuschrecke", "HEU"
    NachtragBIOLK "Hanffaser", "HFU"
    NachtragBIOLK "Hinsch", "HIN"
    NachtragBIOLK "Holle", "HOL"
    NachtragBIOLK "Hoyer", "HOY"
    NachtragBIOLK "Huober", "HUO"
    NachtragBIOLK "Ihle", "IHL"
    NachtragBIOLK "I+M", "IUM"
    NachtragBIOLK "Jatex/Terra Natura", "JAT"
    NachtragBIOLK "Kessel", "KES"
    NachtragBIOLK "Kost Kamm", "KKA"
    NachtragBIOLK "Lakshmi", "LAI"
    NachtragBIOLK "Lavera", "LAV"
    NachtragBIOLK "Lebensbaum, U. Walter", "LEB"
    NachtragBIOLK "Grabower", "LNA"
    NachtragBIOLK "Logona", "Log"
    NachtragBIOLK "Livos", "LVS"
    NachtragBIOLK "Lanwehr", "LWR"
    NachtragBIOLK "Zielke (Savon du M.)", "MID"
    NachtragBIOLK "Mollis", "MLL"
    NachtragBIOLK "MM Cosmetics", "MMC"
    NachtragBIOLK "Moltex", "MOL"
    NachtragBIOLK "Morgenl./ Ege Sun", "MOR"
    NachtragBIOLK "Mr.Evergreen", "MRE"
    NachtragBIOLK "MTC/Maharishi Ayur", "MTC"
    NachtragBIOLK "Neobio", "NEO"
    NachtragBIOLK "Nature Friends", "NFR"
    NachtragBIOLK "Novatex", "NOV"
    NachtragBIOLK "Natracare/Bodyw.", "NTC"
    NachtragBIOLK "Ökonorm", "NWR"
    NachtragBIOLK "Alpro", "PRV"
    NachtragBIOLK "Primavera", "PVL"
    NachtragBIOLK "Radicula", "RAD"
    NachtragBIOLK "Redecker", "RED"
    NachtragBIOLK "Riehm", "RIM"
    NachtragBIOLK "De Rit", "RIT"
    NachtragBIOLK "De Rit", "CHO"
    NachtragBIOLK "De Rit", "MOA"
    NachtragBIOLK "Rosengarten", "ROS"
    NachtragBIOLK "Rösner", "RÖS"
    NachtragBIOLK "Sante", "san"
    NachtragBIOLK "Sanatur", "SAR"
    NachtragBIOLK "Santa Verde", "SAV"
    NachtragBIOLK "Walter Rau/Speick", "SCK"
    NachtragBIOLK "Sekem", "SEF"
    NachtragBIOLK "La Selva", "SEL"
    NachtragBIOLK "Primacos", "SES"
    NachtragBIOLK "sInfo", "SFO"
    NachtragBIOLK "Schöneberger", "SHO"
    NachtragBIOLK "Salus", "SHS"
    NachtragBIOLK "Sonett", "SNT"
    NachtragBIOLK "Sobo Naturkost", "SOB"
    NachtragBIOLK "Sodasan", "SOD"
    NachtragBIOLK "Sonnentor", "STN"
    NachtragBIOLK "Styrums", "STY"
    NachtragBIOLK "Sunval", "SUN"
    NachtragBIOLK "Tapir", "TAP"
    NachtragBIOLK "Tautropfen", "TAU"
    NachtragBIOLK "Terra Soleil", "TES"
    NachtragBIOLK "Tartex", "TRX"
    NachtragBIOLK "999 Energy", "UGH"
    
    NachtragBIOLK "Arche", "URT"
    NachtragBIOLK "Arche", "JOR"
    NachtragBIOLK "Arche", "DAN"
    
    NachtragBIOLK "Vivani/EcoFinia", "VVA"
    NachtragBIOLK "Weleda", "WEL"
    NachtragBIOLK "Werz", "WER"
    NachtragBIOLK "Ecover/Wellments", "WLM"
    NachtragBIOLK "Woodshade", "WOO"
    NachtragBIOLK "Lily Rose", "WWA"
    NachtragBIOLK "Yarrha", "YAR"
    NachtragBIOLK "Zwergenwiese", "ZWE"
    
    NachtragBIOLK "Alnatura", "AAT"
    NachtragBIOLK "Maintal", "ABM"
    NachtragBIOLK "Astrid Heinz, Ra", "AHE"
    NachtragBIOLK "Allos", "ALO"
    NachtragBIOLK "Alva", "ALV"
    NachtragBIOLK "Santa Fe", "AMS"
    NachtragBIOLK "Almawin", "AMW"
    NachtragBIOLK "Apeiron", "APE"
    NachtragBIOLK "Aquabio", "AQU"
    NachtragBIOLK "Arche", "ARC"
    
    'ab hier Pural
    
    NachtragBIOLK "Alma Win", "AMW"
    NachtragBIOLK "Alsan Werk", "ALS"
    NachtragBIOLK "Alva", "ALV"
    NachtragBIOLK "Amazonas", "AMA"
    NachtragBIOLK "Andechser", "AND"
    NachtragBIOLK "Andringa", "ADR"
    NachtragBIOLK "Anis de l Abbaye", "ANI"
    NachtragBIOLK "Aries Umweltprodukte", "ARI"
    NachtragBIOLK "Barnhouse Naturprodukte", "BHO"
    NachtragBIOLK "Bauck", "BAK"
    NachtragBIOLK "Belt s", "BLT"
    NachtragBIOLK "Berchtesgadener Land", "BGL"
    NachtragBIOLK "Bio Kaas Bastiaansen", "BAS"
    NachtragBIOLK "Biokosma", "BKM"
    NachtragBIOLK "Bioland Handelsgesellschaft", "BBW"
    NachtragBIOLK "Bohlsener Mühle", "BOL"
    NachtragBIOLK "Bruno Fischer", "BFS"
    NachtragBIOLK "Byodo Naturkost GmbH", "BYO"
    NachtragBIOLK "Castiglioni", "CGL"
    NachtragBIOLK "Cha Do Teehandels GmbH", "CDO"
    NachtragBIOLK "Chocolat Schönenberger", "COL"
    NachtragBIOLK "Chocoreale De Rit", "CHR"
    NachtragBIOLK "Crudigno", "ORO"
    NachtragBIOLK "Danival Sarl", "DAN"
    NachtragBIOLK "Ecofina L Weidrich GmbH", "VNI"
    NachtragBIOLK "Ecover", "ECV"
    NachtragBIOLK "Erntesegen GmbH", "ERN"
    NachtragBIOLK "Evers Naturkost GmbH", "EVS"
    NachtragBIOLK "Farfalle Essentials AG", "FAR"
    NachtragBIOLK "Flores Farm GmbH", "FLF"
    
    NachtragBIOLK "Fontaine", "FON"
    NachtragBIOLK "Friedrichsdorfer Zwieback", "FZS"
    NachtragBIOLK "FZ Organic Food", "ADR"
    NachtragBIOLK "Gesund und Leben", "GUL"
    NachtragBIOLK "Golden Temple", "GOL"
    NachtragBIOLK "Govinda Natur GmbH", "GOV"
    NachtragBIOLK "GSE Vertrieb GmbH", "GSE"
    NachtragBIOLK "Herbaria", "HER"
    NachtragBIOLK "Holle Baby Food GmbH", "HOL"
    NachtragBIOLK "Isana Naturfeinkost GmbH", "ISA"
    NachtragBIOLK "Kanne Brottrunk", "KAN"
    NachtragBIOLK "Klar", "KLA"
    NachtragBIOLK "La Selva", "SEL"
    NachtragBIOLK "Lauretana", "LTA"
    NachtragBIOLK "Laverana GmbH", "LAV"
    NachtragBIOLK "Lebensbaum U Walter", "LEB"
    NachtragBIOLK "Linea Natura", "LNA"
    NachtragBIOLK "Lima", "LIM"
    NachtragBIOLK "Logona", "Log"
    NachtragBIOLK "Luvos Heilerde", "LUV"
    NachtragBIOLK "Mayka Naturbackwaren GmbH", "MAY"
    NachtragBIOLK "Molenaartje", "MOA"
    NachtragBIOLK "Monte Bianco", "MOB"
    NachtragBIOLK "Morgenland", "MOR"
    NachtragBIOLK "Natracare", "NTC"
    NachtragBIOLK "Natumi", "NTM"
    NachtragBIOLK "Naturana", "NAT"
    NachtragBIOLK "Naturcompagnie", "NCO"
    NachtragBIOLK "Oatly", "OAT"
    NachtragBIOLK "Öma Beer GmbH", "ÖMA"
    NachtragBIOLK "Organix4U", "ORG"
    NachtragBIOLK "Peterstaler", "PMI"
    NachtragBIOLK "Primavera Life GmbH", "PLV"
    NachtragBIOLK "Provamel", "PRV"
    NachtragBIOLK "Pural Vertriebs GmbH", "PUR"
    NachtragBIOLK "Raab Vitalfood GmbH", "RAA"
    NachtragBIOLK "Riegel Wein", "RIE"
    NachtragBIOLK "Rosengarten", "ROS"
    NachtragBIOLK "Runge", "RUN"
    NachtragBIOLK "Savon du Midi", "MID"
    NachtragBIOLK "Schnitzer OHG", "STZ"
    NachtragBIOLK "Sekowa Backtechnik GmbH", "SEK"
    NachtragBIOLK "Sobo Naturkost", "SOB"
    NachtragBIOLK "Sonnentor Kräuterhandelsges", "STN"
    NachtragBIOLK "Taifun", "TAI"
    NachtragBIOLK "Tarpa", "TAR"
    NachtragBIOLK "Tautropfen", "TAU"
    NachtragBIOLK "Terra Sana", "TER"
    NachtragBIOLK "Terra Soleil", "TES"
    NachtragBIOLK "Teutoburger Mühle", "TEU"
    NachtragBIOLK "Tofutown GmbH Viana", "VIA"
    NachtragBIOLK "Topas K Gaiser GmbH", "TOP"
    NachtragBIOLK "Urtekram", "URT"
    NachtragBIOLK "Viana", "VIA"
    NachtragBIOLK "Voelkel Naturgarten", "VOE"
    NachtragBIOLK "Walter Rau", "SCK"
    NachtragBIOLK "Weleda", "WEL"
    NachtragBIOLK "Wellments", "WLM"
    NachtragBIOLK "Werz NaturkornMühle", "WER"
    NachtragBIOLK "Windmill Organics Ltd", "WDM"
    NachtragBIOLK "Zwergenwiese", "ZWE"

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BIOLPURKFuellen"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub NachtragBIOLK(cLiefname As String, cLiefkurz As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    sSQL = "Delete from BIOPURLK where LIEFKURZ = '" & UCase(cLiefkurz) & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into BIOPURLK (LIEFNAME,LIEFKURZ) values ('" & cLiefname & "','" & UCase(cLiefkurz) & "')"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "NachtragBIOLK"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Private Sub Check5_Click()
On Error GoTo LOKAL_ERROR

    If Check5.Value = vbChecked Then
        Frame5.Visible = True
    Else
        Frame5.Visible = False
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check5_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Check8_Click()
On Error GoTo LOKAL_ERROR

    If Check8.Value = vbChecked Then
        zeigWerte Check8.Value, Label2(5), Label2(8), Label2(21), List3, List4, "asc"
    Else
        zeigWerte Check8.Value, Label2(5), Label2(8), Label2(21), List3, List4, "asc"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Check8_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub urzustand1()
    On Error GoTo LOKAL_ERROR

    anzeige "normal", "", Label1(4)
    PRIFlex.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "urzustand1"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Command0_Click(Index As Integer)
On Error GoTo LOKAL_ERROR
    
    Select Case Index
        Case Is = 4
            Text1_KeyUp 4, vbKeyF2, 0
        Case Is = 5
            Text1_KeyUp 2, vbKeyF2, 0
    End Select
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command0_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    Fehlermeldung1
End Sub
Private Sub Command1_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim i As Integer

Select Case Index
    Case 0
        KalkandRund "RUNDEN"
    Case 2 'Zurück
        Frame6.Visible = True
        Frame8.Visible = False
        urzustand1
    Case 1
        If Text2(3).Text <> "" Then
            If IsNumeric(Text2(3).Text) Then
                LEKKalk CDbl(Text2(3).Text)
            End If
        End If
        Text2(3).Text = ""
    Case 4
        gsZSpalte = "EAN"
        gstab = "STADAPRI"
        frmWKL36.Show 1
        'fertig
        
        PricatStep4
    Case 5
        lastvoreinstellungspeichern "EXCELe", frmWKL166, 9
        
        PricatStep5
        PRIFlex.Visible = False
        For i = 0 To 11
            Check1(i).Visible = False
        Next i
        Frame8.Visible = False
        anzeige "normal", "Artikelübernahme erfolgreich beendet.", Label1(4)
    Case 7
        Frame9.Visible = False
        Frame1.Visible = True
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LEKKalk(dABschlag As Double)
    On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim dLEK            As Double
    
    PRIFlex.Redraw = False
    PRIFlex.Row = 0
    For i = 1 To PRIFlex.Rows - 1
        PRIFlex.Row = i
        PRIFlex.Col = SpaltennummerLEKPR
        
        If Not Len(PRIFlex.Text) = 0 Then
            dLEK = PRIFlex.Text
            
            If dLEK <> 0 Then
                PRIFlex.Text = Format((100 - dABschlag) * dLEK / 100, "####0.00")
            End If
        Else
            PRIFlex.Text = "0,00"
        End If
    Next i
    PRIFlex.Refresh
    
    PRIFlex.Redraw = True
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LEKKalk"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub


 









Private Sub KalkandRund(sArt As String)
    On Error GoTo LOKAL_ERROR

    Dim i               As Integer
    Dim dKVKN           As Double
    
    PRIFlex.Redraw = False
    Select Case sArt
        Case "RUNDEN"
    
            PRIFlex.Row = 0
            For i = 1 To PRIFlex.Rows - 1
                PRIFlex.Row = i
                PRIFlex.Col = SpaltennummerKVKNEU
                
                If Not Len(PRIFlex.Text) = 0 Then
                    
                    dKVKN = PRIFlex.Text
                    If dKVKN <> 0 Then
                        PRIFlex.Text = Runden(dKVKN)
                    End If
                Else
                    PRIFlex.Text = "0,00"
                End If
            Next i
            PRIFlex.Refresh
    End Select
    PRIFlex.Redraw = True
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "KalkandRund"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub PricatStep5()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount          As Long
    Dim lAnzRows        As Long
    Dim lAktFeld        As Long
    Dim lUebernommen    As Long
    Dim lAnz            As Long
    Dim lDatum          As Long

    Dim dLEKPR          As Double
    Dim dKVkPr1         As Double
    Dim dKVkPr1Neu      As Double
    Dim dVkPrAlt        As Double
    Dim dVkPrNeu        As Double
    Dim dKVkPrAlt       As Double
    Dim dKVkPrNeu       As Double
    Dim dKarVKPr        As Double

    Dim cPfad           As String
    Dim cSQL            As String
    Dim cKVkPr1         As String
    Dim cKVkPr1NEU      As String
    Dim ctmp            As String
    Dim cFeld           As String
    Dim cWert           As String
    Dim cArtNr          As String
    Dim cAgn            As String
    Dim cPGN            As String
    Dim cLinr           As String
    Dim cLEKPR          As String
    Dim cLiBesNr        As String
    Dim cMinMen         As String
    Dim cMinBest        As String
    Dim cGefuehrt       As String
    Dim cRabatt_OK      As String
    Dim cPreiSchutz     As String
    Dim cGroesse        As String
    Dim cMenge          As String
    Dim cBestand        As String
    
    Dim siAnzeige       As Single
    Dim iRet            As Integer
    Dim bPreisschutz    As Boolean
    Dim rsrs1           As Recordset
    Dim rsRS2           As Recordset
    Dim rsRs3           As Recordset
    Dim rsStadapro      As Recordset
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Die Daten werden in den Stammdatensatz übernommen...", Label1(4)
    
    loeschNEW "STADAPRO", gdBase 'stadapro ist neu
    CreateTable "STADAPRO", gdBase
    
    'auch eine Bestelldatei vorbereiten
    If NewTableSuchenDBKombi("VORSCHLZ", gdBase) = False Then  'das erste Mal
        CreateTable "VORSCHLZ", gdBase
    End If
    
    
    Set rsStadapro = gdBase.OpenRecordset("STADAPRO")
    
    'welche Haken sind gesetzt
    For lcount = 0 To 9
        If Check1(lcount).Value = vbChecked Then
            gbTransfer(lcount + 1) = True
        Else
            gbTransfer(lcount + 1) = False
        End If
    Next
    
    lAnzRows = PRIFlex.Rows
    lAnzRows = lAnzRows - 1
    lAktFeld = 1
    
    If Check1(11).Value = vbChecked Then
    
        cSQL = "Update Artlief inner join importpri on Artlief.artnr = importpri.artnr and Artlief.linr = importpri.linr "
        cSQL = cSQL & " set Artlief.RKZ = 'J' , Artlief.EXDAT = '" & DateValue(Now) & "' "
        gdBase.Execute cSQL, dbFailOnError
    
'       cSQL = "Update Artikel inner join importpri on artikel.artnr = importpri.artnr "
'       cSQL = cSQL & " set Artikel.RKZ = 'J'"  ', artikel.LPZ = 0  "
'       gdBase.Execute cSQL, dbFailOnError
    End If
    
    Label2(35).Caption = lAnzRows - 1
    Label2(35).Refresh
    
    'ab hier geschwindigkeit rausholen
    PRIFlex.Redraw = False  '1.
    picprogress.Visible = True
    txtStatus.Text = 0
    
    lDatum = DateValue(Now)
    
    For lcount = 2 To lAnzRows  'Ab hier wird das ganze Grid abgeklappert
nochmal:
        Label2(33).Caption = lAktFeld
        Label2(33).Refresh
        lAktFeld = lAktFeld + 1
        rsStadapro.AddNew 'Protokoll schreiben ein neuer Datensatz
        rsStadapro!Quelldat = Label2(5).Caption  ' sEinlesedat
        rsStadapro!Datum = lDatum 'DateValue(Now)'2.
        
        PRIFlex.Row = lcount
        
        siAnzeige = siAnzeige + 1
        txtStatus.Text = CStr((100 * siAnzeige) / lAnzRows)
            
        PRIFlex.Col = SpaltennummerArtnr 'Artikelnummer holen
        cArtNr = PRIFlex.Text
        
        'Artikelnummer Check
        
        If Trim(UCase(cArtNr)) = "X" Or cArtNr = "" Then
            If lcount = lAnzRows Then
                Exit For
            ElseIf lcount < lAnzRows Then
                lcount = lcount + 1
                GoTo nochmal
            End If
        End If
        
        PRIFlex.Col = SpaltennummerLINR
        cLinr = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerLIBESNR
        cLiBesNr = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerLEKPR
        cLEKPR = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerKVKPR1
        cKVkPr1 = PRIFlex.Text
        cKVkPr1 = fnMoveComma2Point$(cKVkPr1)
        dKVkPr1 = Val(cKVkPr1)
        
        PRIFlex.Col = SpaltennummerKVKNEU
        cKVkPr1NEU = PRIFlex.Text
        cKVkPr1NEU = fnMoveComma2Point$(cKVkPr1NEU)
        dKVkPr1Neu = Val(cKVkPr1NEU)
        
        PRIFlex.Col = SpaltennummerMINMEN
        cMinMen = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerAGN
        cAgn = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerPGN
        cPGN = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerMINBEST
        cMinBest = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerGEFUEHRT
        cGefuehrt = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerRABATT_OK
        cRabatt_OK = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerPREISSCHU
        cPreiSchutz = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerMENGE
        cMenge = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerBESTAND
        cBestand = PRIFlex.Text
        
        PRIFlex.Col = SpaltennummerVKPR
        If PRIFlex.Text <> "" Then
            dKarVKPr = PRIFlex.Text
        Else
            dKarVKPr = 0
        End If
        
        PRIFlex.Col = SpaltennummerGROESSE
        If PRIFlex.Text <> "" Then
            cGroesse = PRIFlex.Text
        Else
            cGroesse = ""
        End If
        
        cSQL = "Select * from IMPORTPRI where ARTNR = " & cArtNr
        Set rsrs1 = gdBase.OpenRecordset(cSQL)
        
'        If cArtNr = "247997" Then
'            MsgBox "dr"
'        End If
            
        If Not rsrs1.EOF Then
            rsrs1.MoveFirst
            
            cSQL = "Select * from ARTIKEL where ARTNR = " & cArtNr
            Set rsRS2 = gdBase.OpenRecordset(cSQL)
            
            If Not rsRS2.EOF Then
            
                If Not IsNull(rsRS2!PREISSCHU) Then
                    If rsRS2!PREISSCHU = "J" Then
                        bPreisschutz = True
                    Else
                        bPreisschutz = False
                    End If
                Else
                    bPreisschutz = False
                End If
                
                rsRS2.Edit 'ab hier Editmodus in der Artikel
                rsRS2!SYNStatus = "E"
                rsStadapro!Akt = "Änderung"
                
                
                
                rsRS2!RABATT_OK = cRabatt_OK
                rsRS2!PREISSCHU = cPreiSchutz
                rsRS2!GEFUEHRT = cGefuehrt
'                rsRs2!GROESSE = cGroesse
                rsRS2!AWM = "0"
                
                rsStadapro!GEFUEHRT = cGefuehrt
                
                If gbTransfer(3) = True Then
                    rsRS2!MINBEST = cMinBest
                Else
                    rsRS2!MINBEST = 0
                End If
                If Not IsNull(rsRS2!GEFUEHRT) Then
                    ctmp = rsRS2!GEFUEHRT
                Else
                    ctmp = ""
                End If
                If ctmp = "J" Then
                    rsRS2!LASTDATE = lDatum 'DateValue(Now) 3.
                    rsRS2!LASTTIME = TimeValue(Now)
                End If
            Else
            
                cSQL = "Delete from ARTLIEF where ARTNR = " & cArtNr
                gdBase.Execute cSQL, dbFailOnError
                
                rsRS2.AddNew    'ab AddnewModus in der Artikel
                rsRS2!AUFDAT = DateValue(Now)
                rsRS2!SYNStatus = "A"
                rsStadapro!Akt = "Neuheit"
                
                If gbTransfer(5) = True Then
                    rsRS2!GEFUEHRT = "J"
                Else
                    rsRS2!GEFUEHRT = "N" 'cGefuehrt
                    rsStadapro!GEFUEHRT = rsRS2!GEFUEHRT
                End If
                
                If gbTransfer(3) = True Then
                    rsRS2!MINBEST = cMinBest
                Else
                    rsRS2!MINBEST = 0
                End If
                
                
                rsRS2!RABATT_OK = cRabatt_OK
'                rsRs2!GROESSE = cGroesse
                rsRS2!BONUS_OK = "J"
                rsRS2!UMS_OK = "J"
                rsRS2!PREISSCHU = cPreiSchutz
                rsRS2!LASTDATE = lDatum 'DateValue(Now)4.
                rsRS2!LASTTIME = TimeValue(Now)
                rsRS2!ekpr = rsrs1!lekpr
                rsRS2!ETIMERK = "N"
                rsRS2!AWM = "98"
                
            End If
            
            If Check1(10).Value = vbChecked Then
            
                cSQL = "Update Artlief set SYNSTATUS = 'D' where artnr = " & rsrs1!artnr & " "
                gdBase.Execute cSQL, dbFailOnError
            End If
            cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLinr & " and (SYNSTATUS <> 'D' or SYNSTATUS is null)"
'            cSQL = "Select * from ARTLIEF where (SYNSTATUS = 'E' or SYNSTATUS = 'A' or SYNSTATUS is null) and ARTNR = " & cArtNr & " and LINR = " & cLinr
            Set rsRs3 = gdBase.OpenRecordset(cSQL)
            
            If Not IsNull(rsrs1!artnr) Then
                rsStadapro!artnr = rsrs1!artnr
            Else
                rsStadapro!artnr = ""
            End If
            rsRS2!artnr = rsrs1!artnr
                     
            If Not IsNull(rsrs1!BEZEICH) Then
                rsStadapro!BEZEICH = rsrs1!BEZEICH
            Else
                rsStadapro!BEZEICH = ""
            End If
            
            If rsStadapro!Akt = "Neuheit" Then
                rsRS2!BEZEICH = rsrs1!BEZEICH
            Else
                If (gbTransfer(10) = True) Then
                    rsRS2!BEZEICH = rsrs1!BEZEICH
                Else
                    rsRS2!BEZEICH = rsRS2!BEZEICH
                End If
            End If
        
            rsRS2!linr = rsrs1!linr
            
            If rsStadapro!Akt = "Neuheit" Then
                rsRS2!LPZ = rsrs1!LPZ
            Else
                If gbTransfer(8) = True Then
                    rsRS2!LPZ = rsrs1!LPZ
                Else
                    rsRS2!LPZ = rsRS2!LPZ
                End If
            End If
                              
            
            rsRS2!LIBESNR = rsrs1!LIBESNR
                               
            If Not IsNull(rsRS2!lekpr) Then
                dLEKPR = rsRS2!lekpr
            Else
                dLEKPR = 0
            End If
            cWert = Format$(dLEKPR, "####0.00")
            rsStadapro!LEK_ALT = cWert

            If Not IsNull(rsrs1!lekpr) Then
                dLEKPR = rsrs1!lekpr
            Else
                dLEKPR = 0
            End If
            cWert = Format$(dLEKPR, "####0.00")

            If gbTransfer(7) = True Then
                rsStadapro!LEK_NEW = cWert

                rsRS2!lekpr = rsrs1!lekpr
            Else
                rsStadapro!LEK_NEW = rsStadapro!LEK_ALT

                rsRS2!lekpr = rsRS2!lekpr
            End If
                   
            If rsRS2!KVKPR1 <> dKVkPr1Neu Then
                If PRIFlex.CellBackColor = vbYellow Then
                    rsStadapro!AUTOKALK = True
                End If
                
                '** Bei Preisänderung ETIDRU füllen **
                If rsStadapro!Akt = "Änderung" Then
                    rsStadapro!Akt = "Preisänderung"
                End If
                
                If Not bPreisschutz Then
                    SchreibeEtiDruWKL11 rsrs1, rsRS2, dKVkPr1Neu, cArtNr
                End If
                
            End If
                        
            If Not IsNull(rsRS2!vkpr) Then
                dVkPrAlt = rsRS2!vkpr
            Else
                dVkPrAlt = 0
            End If
            cWert = Format$(dVkPrAlt, "####0.00")
            rsStadapro!VKPR_ALT = cWert

            If Not IsNull(rsRS2!KVKPR1) Then
                dKVkPrAlt = rsRS2!KVKPR1
            Else
                dKVkPrAlt = 0
            End If
            cWert = Format$(dKVkPrAlt, "####0.00")
            rsStadapro!KVK_ALT = cWert

            If Not IsNull(rsrs1!vkpr) Then
                dVkPrNeu = rsrs1!vkpr
            Else
                dVkPrNeu = 0
            End If
            
            cWert = Format$(dKarVKPr, "######0.00")
            If gbTransfer(6) = False Then
                rsStadapro!VKPR_NEW = cWert
            Else
                cWert = Format$(dVkPrNeu, "######0.00")
                rsStadapro!VKPR_NEW = cWert
                rsRS2!vkpr = dKarVKPr
            End If
            
            
            If rsStadapro!Akt = "Neuheit" Then
                If (gbTransfer(1) = True) Then
                    rsRS2!KVKPR1 = dKVkPr1Neu
                    rsRS2!vkpr = dKarVKPr
                    cWert = Format$(dKVkPr1Neu, "####0.00")
                    rsStadapro!KVK_NEW = cWert
                Else
                    rsRS2!KVKPR1 = dKVkPr1Neu
                    rsStadapro!KVK_NEW = dKVkPr1Neu
                End If
            Else
                If (gbTransfer(1) = True) And Not bPreisschutz Then
                    
                    rsRS2!KVKPR1 = dKVkPr1Neu
                    cWert = Format$(dKVkPr1Neu, "####0.00")
                    rsStadapro!KVK_NEW = cWert
                Else
                    If dKVkPrAlt > 0 Then
                        ctmp = Format$(dKVkPrAlt, "####0.00")
                    Else
                        ctmp = Format$(dKarVKPr, "####0.00")
                    End If
                    rsRS2!KVKPR1 = ctmp
                    rsStadapro!KVK_NEW = ctmp
                End If
            End If
            
            If rsStadapro!Akt = "Neuheit" Then
                If IsNumeric(cAgn) Then
                    rsRS2!AGN = cAgn
                Else
                    rsRS2!AGN = 0
                End If
            Else
                If gbTransfer(9) = True Then
                    If IsNumeric(cAgn) Then
                        rsRS2!AGN = cAgn
                    Else
                        rsRS2!AGN = 0
                    End If
                Else
                    rsRS2!AGN = rsRS2!AGN
                End If
            End If
            
''            'RKZ check
''            rsRs2!RKZ = rsrs1!RKZ
''            If rsRs2!RKZ = "J" Then
''                If Not IsNull(rsRs2!EXDAT) Then
''                    If CLng(rsRs2!EXDAT) = 0 Then
''                        rsRs2!EXDAT = DateValue(Now)
''                    End If
''                Else
''                    rsRs2!EXDAT = DateValue(Now)
''                End If
''            Else
''                rsRs2!EXDAT = 0
''            End If
            
            If rsrs1!EAN <> rsRS2!EAN Then
                If Not IsNull(rsRS2!EAN2) Then
                    rsRS2!EAN3 = rsRS2!EAN2
                End If
                rsRS2!EAN2 = rsRS2!EAN
            End If
            rsRS2!EAN = rsrs1!EAN
            
            If gbTransfer(2) = True Then rsRS2!MINMEN = rsrs1!MINMEN
            If rsStadapro!Akt = "Neuheit" And (gbTransfer(2) = False) Then rsRS2!MINMEN = 1
        
            rsRS2!MWST = rsrs1!MWST
        
            If gbTransfer(4) = True Or (rsStadapro!Akt = "Neuheit") Then
                rsRS2!NOTIZEN = Left(rsrs1!NOTIZEN, 25)
            End If
            
            rsRS2!GROESSE = rsrs1!GROESSE
            rsRS2!FARBNR = rsrs1!FARBNR
        
            rsRS2!INHALT = rsrs1!INHALT
            rsRS2!INHALTBEZ = rsrs1!INHALTBEZ
            rsRS2!GRUNDPREIS = rsrs1!GRUNDPREIS
            rsRS2!BESTAND = rsrs1!BESTAND
            rsRS2.Update
            
            If rsRs3.EOF Then
                rsRs3.AddNew
                rsRs3!SYNStatus = "A"
            Else
                rsRs3.Edit
                rsRs3!SYNStatus = "E"
            End If
            rsRs3!artnr = Val(cArtNr)
            rsRs3!linr = Val(cLinr)
            rsRs3!LIBESNR = cLiBesNr
            
            'Check mal den Listenek fürs Protokoll
            
            If Not IsNull(rsRs3!lekpr) Then
            
                If Format(cLEKPR, "####0.00") = Format(rsRs3!lekpr, "####0.00") Then
                
                Else
                    rsStadapro!Akt = "Preisänderung"
                    rsStadapro!LEK_ALT = rsRs3!lekpr
                    rsStadapro!LEK_NEW = Format(cLEKPR, "####0.00")
                
                End If
            End If
            
            
            rsRs3!lekpr = Format(cLEKPR, "####0.00")
            rsRs3!MINMEN = Val(cMinMen)
            
            
            'RKZ check
            rsRs3!RKZ = rsrs1!RKZ
            If rsRs3!RKZ = "J" Then
                If Not IsNull(rsRs3!EXDAT) Then
                    If CLng(rsRs3!EXDAT) = 0 Then
                        rsRs3!EXDAT = DateValue(Now)
                    End If
                Else
                    rsRs3!EXDAT = DateValue(Now)
                End If
            Else
                rsRs3!EXDAT = 0
            End If
            
            
            
            rsRs3.Update
            
        End If
        rsStadapro.Update
        
        If Val(cMenge) > 0 Then
            InsertinBestellDatei cArtNr, Val(cMenge), cLinr
        End If
    Next lcount
    picprogress.Visible = False
    
    rsRs3.Close
    rsRS2.Close
    rsrs1.Close
    rsStadapro.Close
    
    If Datendrin("VORSCHLZ", gdBase) = True Then
        BestellungWegSpeichern cLinr
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PricatStep5"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
End Sub
Private Function BestellungWegSpeichern(cLinr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cBezeich    As String
    Dim lRet        As Long
    Dim ctmp        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim lAnzTable   As Long
    Dim lcount      As Long
    Dim bgefunden   As Boolean
    Dim cZiel       As String
    Dim lAktTable   As Long
    Dim lDatum      As Long
    Dim cDatum      As String
    Dim iStufe      As Integer
    

    'Namen zuweisen
    lcount = 65
    bgefunden = True
    Do While bgefunden
        cZiel = "Q" & cLinr & Chr$(lcount)
        If NewTableSuchenDBKombi("Q" & cLinr & Chr$(lcount), gdBase) Then
            bgefunden = True
            lcount = lcount + 1
        Else
            bgefunden = False
        End If
    Loop
    iStufe = 5
    
    If lcount > 89 Then
        Dim ctempa As String
        ctempa = "Bestellvorschlag speichern nicht möglich. Die Vergabe eines Dateinamens ist gescheitert." & vbCrLf
        ctempa = ctempa & "Löschen Sie erledigte Bestellungen im Wareneingang aus Bestellung!"
        MsgBox ctempa, vbOKOnly + vbInformation, "Winkiss Hinweis:"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    
    loeschNEW cZiel, gdBase
    
    cSQL = "Select * into " & cZiel & " from VORSCHLZ "
    cSQL = cSQL & " where BESTVOR > 0 "
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Delete from TABDATUM where TABNAME like '" & cZiel & "*' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into TABDATUM (Tabname,Tabdate) values"
    cSQL = cSQL & " ( '" & cZiel & "','" & DateValue(Now) & "')"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from BESTREST where DATEINAME = '" & cZiel & ".DBF'"
    gdBase.Execute cSQL, dbFailOnError
    
'            SetzeNeuheitenwiederaufgrau gsLinr
    
'            If Text12.Text <> "" Then
'                If IsNumeric(Text12.Text) Then
'                    schreibeAufnr CLng(Text12.Text), cZiel
'                End If
'            End If
    
    lDatum = Fix(Now)
    cDatum = Trim$(Str$(lDatum))
    
    cSQL = "Insert into BESTREST "
    cSQL = cSQL & "Select " & cLinr & " as LINR, "
    cSQL = cSQL & "ARTNR, LEKPR, BESTVOR, '" & cZiel & ".DBF' as DATEINAME, "
    cSQL = cSQL & cDatum & " as BEST_DATUM, " & cDatum & " as UPD_DATUM "
    cSQL = cSQL & " from " & cZiel & " where BESTVOR <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    'alle Artikel, die bestellt werden - auf geführt = 'J' setzen
    cSQL = "Update Artikel inner join " & cZiel & " on Artikel.artnr = " & cZiel & ".artnr  "
    cSQL = cSQL & " Set Artikel.gefuehrt = 'J' "
    cSQL = cSQL & " where " & cZiel & ".BESTVOR <> 0 "
    gdBase.Execute cSQL, dbFailOnError
    'Ende, alle Artikel, die bestellt werden - auf geführt = 'J' setzen
    
    MsgBox "Lieferavis als Bestellung " & cZiel & " gespeichert!", vbInformation, "Winkiss Hinweis:"
    
    loeschNEW "VORSCHLZ", gdBase
                
    Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "BestellungWegSpeichern"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub InsertinBestellDatei(cArtNr As String, lMenge As Long, cLinr As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    sSQL = "insert into vorschlz "
    sSQL = sSQL & " Select distinct"
    sSQL = sSQL & "  A.ARTNR "
    sSQL = sSQL & ", A.BEZEICH "
    sSQL = sSQL & ", A.AGN "
    sSQL = sSQL & ", A.PGN "
    sSQL = sSQL & ", B.LEKPR "
    sSQL = sSQL & ", A.AWM "
    sSQL = sSQL & ", A.EKPR "
    sSQL = sSQL & ", A.VKPR "
    sSQL = sSQL & ", " & cLinr & " as LINR "
    sSQL = sSQL & ", " & lMenge & " as BESTVOR "
    sSQL = sSQL & ", B.LIBESNR "
    sSQL = sSQL & ", A.KVKPR1 "
    sSQL = sSQL & ", A.EAN "
    sSQL = sSQL & ", 0 as MOPREIS "
    sSQL = sSQL & ", B.RKZ "
    sSQL = sSQL & ", A.LPZ "
    sSQL = sSQL & ", A.NOTIZEN "
    sSQL = sSQL & ", B.MINMEN "
    sSQL = sSQL & ", 0 as MINBEST "
    sSQL = sSQL & ", 0 as INBEST "
    sSQL = sSQL & ", 0 as ANZEIGE "
    sSQL = sSQL & ", 0 as FAKTOR "
    sSQL = sSQL & ", A.BESTAND "
    sSQL = sSQL & ", NULL as LPZ_BIS "
    sSQL = sSQL & ", NULL as LPZ_VON "
    sSQL = sSQL & ", 0 as BEVORRAT "
    sSQL = sSQL & ", 0 as EINDECK "
    sSQL = sSQL & ", NULL as VKAMo1 "
    sSQL = sSQL & ", NULL as VKVMo1 "
    sSQL = sSQL & ", NULL as VKLJ1 "
    sSQL = sSQL & ", NULL as VKVJ1 "
    sSQL = sSQL & ", NULL as MITTEILUNG"
    sSQL = sSQL & ", NULL as LJ1 , NULL as LJ2, NULL as LJ3, NULL as LJ4, NULL as LJ5,NULL as LJ6, NULL as LJ7,NULL as LJ8,NULL as LJ9,NULL as LJ10,NULL as LJ11,NULL as LJ12 "
    sSQL = sSQL & ", NULL as VJ1 , NULL as VJ2, NULL as VJ3, NULL as VJ4, NULL as VJ5,NULL as VJ6, NULL as VJ7,NULL as VJ8,NULL as VJ9,NULL as VJ10,NULL as VJ11,NULL as VJ12 "
    sSQL = sSQL & " from ARTIKEL A, ARTLIEF B "
    sSQL = sSQL & " where B.LINR = " & cLinr & " "
    sSQL = sSQL & " and A.ARTNR = " & cArtNr & " "
    sSQL = sSQL & " and A.ARTNR = B.ARTNR "
    sSQL = sSQL & " and A.GEFUEHRT = 'J'"
    sSQL = sSQL & " and (A.SYNSTATUS = 'E' or A.SYNSTATUS = 'A' or A.SYNSTATUS is null)"
    sSQL = sSQL & " order by A.LPZ, A.BEZEICH "
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "InsertinBestellDatei"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub Command4_Click(Index As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim btxt As Boolean
    Dim iFileNr As Integer
    Dim lRet As Long
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Select Case Index
        Case 0 To 2
        
            If Datendrin("Stadapro", gdBase) = True Then
                Select Case Index
                    Case Is = 0 'alle Artikel hier auf geführtcheck achten
                    
                        If CG.Value Then
                            reportbildschirm "dWKL11a", "aWKL11e"
                        Else
                            reportbildschirm "dWKL11a", "aWKL11a"
                        End If
                    Case Is = 1 'neue Artikel
                        reportbildschirm "dWKL11b", "aWKL11b"
                    Case Is = 2 'Preisänderungen hier auf geführtcheck achten
                    
                        If CG.Value Then
                            reportbildschirm "dWKL11a", "aWKL11f"
                        Else
                            reportbildschirm "dWKL11a", "aWKL11c"
                        End If
                End Select
            Else
                anzeige "rot", "keine Protokoll - Daten vorhanden", Label1(4)
            End If
    End Select
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command4_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Private Sub Command5_Click(Index As Integer)
On Error GoTo LOKAL_ERROR

Dim cVorauswahlspalte As String

Select Case Index
    Case 0
        Unload frmWKL166
    Case 1
        gsHelpstring = "Excel Import"
        frmWKL110.Show 1
    
    Case 2
        Label3(0).Caption = ""
        Label3(1).Caption = ""
        Label3(2).Caption = ""
        
    Case 3
        pfadseekExcel
    Case 4
        Frame2.Visible = False
    Case 5
        
        If checkZuordnung(Label2(27).Caption & " " & Label2(30).Caption) Then
            'weiter aktivieren
            
            Frame7.Visible = True
            Frame2.Visible = False
            
            Zuordnunganzeigen List13, Label2(27).Caption & " " & Label2(30).Caption
        
            Command5(20).Enabled = True
            Command5(20).BackColor = vbRed
        Else
        
            Label3(0).Caption = ""
            Label3(1).Caption = ""
            Label3(2).Caption = ""
            Frame3.Visible = True
            Frame2.Visible = False
            
            Vorauswahlcheck Label2(5).Caption & " " & Label2(8).Caption, Check2, Label2(28)
            
            Zuordnunganzeigen List12, Label2(5).Caption & " " & Label2(8).Caption
            füllelist5
            LoeschInlist5 List12
        End If
    Case 6
        Frame2.Visible = True
    Case 7
        cVorauswahlspalte = ermVorauswahl(Label2(5).Caption & " " & Label2(8).Caption)
        
        If cVorauswahlspalte = "" Then
        
            anzeige "rot", "Benennen Sie erst eine Vorauswahl!", Label1(4)
            Check2.BackColor = vbRed
            Exit Sub
        Else
            Frame6.Visible = True
            Frame3.Visible = False
            
            erstellegrid cVorauswahlspalte
            zeigVorauswahlWerte Label2(5), Label2(8), cVorauswahlspalte, False
            
            Text1(4).Text = ZeigeEtwasforExcelimport(Label2(5).Caption & " " & Label2(8).Caption, "AGN")
            Text1(2).Text = ZeigeEtwasforExcelimport(Label2(5).Caption & " " & Label2(8).Caption, "LINR")
            Text1(0).Text = ZeigeEtwasforExcelimportalsString(Label2(5).Caption & " " & Label2(8).Caption, "KUERZEL")
            Text2(0).Text = ZeigeEtwasforExcelimportalsString(Label2(5).Caption & " " & Label2(8).Caption, "MWSTV")
            Text2(1).Text = ZeigeEtwasforExcelimportalsString(Label2(5).Caption & " " & Label2(8).Caption, "MWSTE")
            If Text2(0).Text <> "" Or Text2(1).Text <> "" Then
                Check5.Value = vbChecked
            Else
                Check5.Value = vbUnchecked
            End If
        End If
    Case 8
        Frame2.Visible = True
        Frame3.Visible = False
    Case 9
        Frame1.Visible = False
        Frame9.Visible = True
    Case 10
    
        Label3(0).Caption = ""
        Label3(1).Caption = ""
        Label3(2).Caption = ""
        Frame3.Visible = True
        Frame4.Visible = False
    Case 11 'Zuordnen
        Zuordnung Label2(5).Caption & " " & Label2(8).Caption, 0
        Zuordnunganzeigen List12, Label2(5).Caption & " " & Label2(8).Caption
        If checkZuordnung(Label2(5).Caption & " " & Label2(8).Caption) Then
            'weiter aktivieren
            Command5(7).Enabled = True
            Command5(7).BackColor = vbRed
        Else
            Command5(7).Enabled = False
            Command5(7).BackColor = Command5(8).BackColor
        End If
    
    Case 15
        Label3(0).Caption = ""
        Label3(1).Caption = ""
        Label3(2).Caption = ""
        Frame3.Visible = True
        Frame7.Visible = False
            
        Vorauswahlcheck Label2(5).Caption & " " & Label2(8).Caption, Check2, Label2(28)
        Zuordnungloeschen Label2(5).Caption & " " & Label2(8).Caption
        Check2.Value = vbUnchecked
        Check2.Caption = ""
        
        füllelist5
        List12.Clear
        Command5(7).Enabled = False
        Command5(7).BackColor = Command5(8).BackColor
    Case 16
        Zuordnungloeschen Label2(5).Caption & " " & Label2(8).Caption
        füllelist5
        List12.Clear
        Command5(7).Enabled = False
        Command5(7).BackColor = Command5(8).BackColor
    Case 17 'von 6 auf 7 oder 3
        If checkZuordnung(Label2(27).Caption & " " & Label2(30).Caption) Then
            'weiter aktivieren
            
            Frame7.Visible = True
            Frame6.Visible = False
            
            Zuordnunganzeigen List13, Label2(27).Caption & " " & Label2(30).Caption
        
            Command5(20).Enabled = True
            Command5(20).BackColor = vbRed
        Else
            Label3(0).Caption = ""
            Label3(1).Caption = ""
            Label3(2).Caption = ""
            Frame3.Visible = True
            Frame6.Visible = False
            
            Command5(7).Enabled = False
            Command5(7).BackColor = Command5(8).BackColor
            
            Vorauswahlcheck Label2(5).Caption & " " & Label2(8).Caption, Check2, Label2(28)
            
            Zuordnunganzeigen List12, Label2(5).Caption & " " & Label2(8).Caption
            füllelist5
            LoeschInlist5 List13
        End If
    Case 18
    
        If Not gueltigeLINR(Val(Text1(2).Text)) Then
            anzeige "rot", "Geben Sie bitte eine gültige Lieferantennummer an!", Label1(4)
            Text1(2).SetFocus
            Exit Sub
        End If
        
        If Not gueltigeAGN(Val(Text1(4).Text)) Then
            anzeige "rot", "Geben Sie bitte eine gültige Artikelgruppennummer an!", Label1(4)
            Text1(4).SetFocus
            Exit Sub
        End If
        
        cVorauswahlspalte = ermVorauswahl(Label2(5).Caption & " " & Label2(8).Caption)
        picprogress.Visible = True
        If Not EXCELStep3(Label2(5).Caption, Label2(8).Caption, Val(Text1(2).Text), Val(Text1(4).Text), Text1(0).Text, cVorauswahlspalte) Then Exit Sub
        
        Frame8.Visible = True
        Frame6.Visible = False
        Me.Refresh
        PricatStep4
        
    Case 19 'von 7 auf 2
        Frame2.Visible = True
        Frame7.Visible = False
    Case 20 'von 7 auf 6
        cVorauswahlspalte = ermVorauswahl(Label2(5).Caption & " " & Label2(8).Caption)
        
        If cVorauswahlspalte = "" Then
        
        Else
            Frame6.Visible = True
            Frame7.Visible = False
            
            erstellegrid cVorauswahlspalte
            zeigVorauswahlWerte Label2(5), Label2(8), cVorauswahlspalte, False
            
            Text1(4).Text = ZeigeEtwasforExcelimport(Label2(5).Caption & " " & Label2(8).Caption, "AGN")
            Text1(2).Text = ZeigeEtwasforExcelimport(Label2(5).Caption & " " & Label2(8).Caption, "LINR")
            Text1(0).Text = ZeigeEtwasforExcelimportalsString(Label2(5).Caption & " " & Label2(8).Caption, "KUERZEL")
            Text2(0).Text = ZeigeEtwasforExcelimportalsString(Label2(5).Caption & " " & Label2(8).Caption, "MWSTV")
            Text2(1).Text = ZeigeEtwasforExcelimportalsString(Label2(5).Caption & " " & Label2(8).Caption, "MWSTE")
            If Text2(0).Text <> "" Or Text2(1).Text <> "" Then
                Check5.Value = vbChecked
            Else
                Check5.Value = vbUnchecked
            End If
        End If
    Case 21
        flex "entfernen"
        Command5(18).Enabled = False
        Command5(18).BackColor = Command5(17).BackColor
    Case 22
        flex "auswählen"
        Command5(18).Enabled = True
        Command5(18).BackColor = vbRed
        
End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ZeigeEtwasforExcelimport(sSchemaname As String, sEtwas As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    ZeigeEtwasforExcelimport = 0

    sSQL = "Select " & sEtwas & " as Etwas from spzuordlinr where schemaname = '" & sSchemaname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Etwas) Then
            ZeigeEtwasforExcelimport = rsrs!Etwas
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeEtwasforExcelimport"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Private Function ZeigeEtwasforExcelimportalsString(sSchemaname As String, sEtwas As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    ZeigeEtwasforExcelimportalsString = ""

    sSQL = "Select " & sEtwas & " as Etwas from spzuordlinr where schemaname = '" & sSchemaname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Etwas) Then
            ZeigeEtwasforExcelimportalsString = rsrs!Etwas
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ZeigeEtwasforExcelimportalsString"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Private Sub speicherLINRandAGNforExcelimport(sSchemaname As String, lLinr As Long, lagn As Long, cLiefkurz As String, sMWSTV As String, sMWSTE As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    'Delete
    sSQL = "Delete from SPZUORDLINR where Schemaname = '" & sSchemaname & "'"
    gdBase.Execute sSQL, dbFailOnError

    'Insert
    sSQL = "Insert into SPZUORDLINR (LINR,AGN,SCHEMANAME,Kuerzel,MWSTV,MWSTE)" '
    sSQL = sSQL & " values "
    sSQL = sSQL & "(" & lLinr & " "
    sSQL = sSQL & "," & lagn & " "
    sSQL = sSQL & ",'" & sSchemaname & "'"
    sSQL = sSQL & ",'" & cLiefkurz & "'"
    sSQL = sSQL & ",'" & sMWSTV & "'"
    sSQL = sSQL & ",'" & sMWSTE & "'"
    sSQL = sSQL & ")"
    gdBase.Execute sSQL, dbFailOnError
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherLINRandAGNforExcelimport"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Function EXCELStep3(sdat As String, sTabBlatt As String, lLinr As Long, lagn As Long, cLiefkurz As String, cVorauswahlspalte As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim cPfad1      As String
    Dim oldpath     As String
    Dim newpath1    As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lAnz        As Long
    Dim i           As Integer
    Dim j           As Integer
    Dim lfnr1       As Long
    Dim gsExcel50   As String
    Dim dbExcel     As Database
    Dim rsZ         As Recordset
    Dim cVKPR       As String
    Dim cLEKPR      As String
    Dim cMWST       As String
    Dim siAnzeige   As Single
    Dim lcount      As Long
    Dim bOr         As Boolean
    Dim lCountGew   As Long
    Dim lLinie      As Long
    Dim sMWSTV      As String
    Dim sMWSTE      As String
    
    Dim cEAN        As String
    Dim cLiBesNr    As String
    Dim lBIOAGN     As Long
    
    Dim cFeld       As String
    
    bOr = False

    EXCELStep3 = False
    
    cPfad1 = gcDBPfad    'Dabapfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    'vorbereitung der Importtabelle
    '2. dann erstellen
    loeschNEW "meister", gdApp
    CreateTable "MEISTER", gdApp
    
    SpalteAnfuegenNEW "MEISTER", "MENGE", "LONG", gdApp
    
    '2a. Datei kopieren in den Anwendungspfad+Stammda
    oldpath = sExcelpfad & "\" & sdat
    newpath1 = cPfad1 & "Stammda\EXCEL\" & sdat
    lRet = CopyFile(oldpath, newpath1, lfail)

    Pause (1)
    txtStatus.Text = "0"
    
    '5. Datensätze werden eingelesen
    
    lfnr1 = 0
    SchemaZuordnung sdat & " " & sTabBlatt
    
    If Check5.Value = vbChecked Then
        sMWSTV = Text2(0).Text
        sMWSTE = Text2(1).Text
    Else
        sMWSTV = ""
        sMWSTE = ""
    End If
    
    speicherLINRandAGNforExcelimport Label2(5).Caption & " " & Label2(8).Caption, lLinr, lagn, cLiefkurz, sMWSTV, sMWSTE
    
    Set rsrs = gdApp.OpenRecordset("MEISTER")
    gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
    
    Set dbExcel = OpenDatabase(sExcelpfad & "\" & sdat, 0, 0, gsExcel50)
    
    
    Dim tdTd As TableDef
    Dim lAnzFelder As Long
    Dim lTyp As Long
    Dim cFeldTyp As String
    
    Set tdTd = dbExcel.TableDefs(sTabBlatt)

    lAnzFelder = tdTd.Fields.Count
    For lcount = 0 To lAnzFelder - 1
        If UCase(cVorauswahlspalte) = UCase(tdTd.Fields(lcount).name) Then
            lTyp = tdTd.Fields(lcount).Type
            
            Select Case lTyp
                Case Is = dbDate
                    cFeldTyp = "Datum"
                Case Is = dbText
                    cFeldTyp = "Text"
                Case Is = dbMemo
                    cFeldTyp = "Memofeld"
                Case Is = dbBoolean
                    cFeldTyp = "Ja/Nein-Schalter"
                Case Is = dbInteger
                    cFeldTyp = "Ganzzahl"
                Case Is = dbLong
                    cFeldTyp = "Ganzzahl"
                Case Is = dbCurrency
                    cFeldTyp = "Währung"
                Case Is = dbSingle
                    cFeldTyp = "Kommazahl"
                Case Is = dbDouble
                    cFeldTyp = "Kommazahl"
                Case Is = dbByte
                    cFeldTyp = "Byte"
                Case Is = dbLongBinary
                    cFeldTyp = "OLE-Objekt"
            End Select
        End If
    Next lcount
    
    sSQL = "Select "
    sSQL = sSQL & "  " & sFremdSpalteEX(0) & " as " & sKissSpalteEX(0) & " "
    For i = 1 To byAnzahlSpaltenEX - 1
        sSQL = sSQL & "  ," & sFremdSpalteEX(i) & " as " & sKissSpalteEX(i) & " "
    Next i
    sSQL = sSQL & "   from [" & sTabBlatt & "] "
    
    MSFlexGrid1.Redraw = False
    lCountGew = 0
    For lcount = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Row = lcount

        If MSFlexGrid1.Text = "ausgewählt" Then
            lCountGew = lCountGew + 1
        End If
    Next lcount
    
    For lcount = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Row = lcount
        
        If MSFlexGrid1.Text = "ausgewählt" Then
            MSFlexGrid1.Col = 0
            
            
            If bOr = True Then
                If cFeldTyp = "Text" Then
                    sSQL = sSQL & " or [" & cVorauswahlspalte & "] =  '" & MSFlexGrid1.Text & "' "
                ElseIf cFeldTyp = "Kommazahl" Then
                    sSQL = sSQL & " or  [" & cVorauswahlspalte & "] =  " & Val(MSFlexGrid1.Text) & " "
                Else
                    sSQL = sSQL & " or cstr([" & cVorauswahlspalte & "]) =  '" & MSFlexGrid1.Text & "' "
                End If
            Else
                If lCountGew = 1 Then
                    If cFeldTyp = "Kommazahl" Then
                        sSQL = sSQL & " where [" & cVorauswahlspalte & "] =  " & MSFlexGrid1.Text & " "
                    Else
                        sSQL = sSQL & " where [" & cVorauswahlspalte & "] =  '" & MSFlexGrid1.Text & "' "
                    End If
                ElseIf lCountGew > 1 Then
                    
                    If cFeldTyp = "Text" Then
                        sSQL = sSQL & " where ( [" & cVorauswahlspalte & "] =  '" & MSFlexGrid1.Text & "' "
                    ElseIf cFeldTyp = "Kommazahl" Then
                        sSQL = sSQL & " where ( [" & cVorauswahlspalte & "] =  " & Val(MSFlexGrid1.Text) & " "
                    Else
                        sSQL = sSQL & " where ( cstr([" & cVorauswahlspalte & "]) =  '" & MSFlexGrid1.Text & "' "
                    End If
                    
    
                End If
            End If
        
            bOr = True
            
        End If
    Next lcount
    
    If lCountGew = 1 Then

    ElseIf lCountGew > 1 Then
        sSQL = sSQL & " ) "
    End If

'    sSQL = sSQL & " and not " & sFremdSpalteEX(0) & " is null "

    MSFlexGrid1.Redraw = True
    
'    MsgBox sSQL
    Set rsZ = dbExcel.OpenRecordset(sSQL)
    If Not rsZ.EOF Then
        rsZ.MoveLast
        lAnz = rsZ.RecordCount
        rsZ.MoveFirst
        Do While Not rsZ.EOF
            rsrs.AddNew
            
            rsrs!AGN = lagn
            rsrs!LPZ = 1
            
            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)
            
            lfnr1 = lfnr1 + 1
            rsrs!lfnr = lfnr1
                
            If Not IsNull(rsZ!EANEX) Then
            
                cEAN = Trim(rsZ!EANEX)
                cEAN = SwapStr(cEAN, " ", "")
                cEAN = Val(cEAN)
                
                If Len(cEAN) = 11 Then
                    cEAN = "0" & cEAN
                End If
                
                If Len(cEAN) > 13 Then
                    rsrs!EAN = ""
                Else
                    rsrs!EAN = cEAN
                End If
            End If

            If Not IsNull(rsZ!LIBESNREX) Then
                cLiBesNr = Trim(rsZ!LIBESNREX)
                cLiBesNr = SwapStr(cLiBesNr, " ", "")
                
                If Len(cLiBesNr) > 13 Then
                    rsrs!LIBESNR = ""
                Else
                    rsrs!LIBESNR = cLiBesNr
                End If
            End If
            
            If Not IsNull(rsZ!VKPREX) Then
                cVKPR = rsZ!VKPREX
                
                If Right(cVKPR, 1) = "€" Then
                    rsrs!vkpr = Left(cVKPR, Len(cVKPR) - 1)
                Else
                    rsrs!vkpr = rsZ!VKPREX
                End If
            End If
            
            If Not IsNull(rsZ!BEZEICHEX) Then
                If cLiefkurz <> "" Then
                    cFeld = Trim(cLiefkurz & " " & CStr(rsZ!BEZEICHEX))
                    For j = 20 To 2 Step -1
                        cFeld = SwapStr(cFeld, Space(j), " ")
                    Next j
                    rsrs!BEZEICH = Left(cFeld, 35)
                Else
                    cFeld = CStr(rsZ!BEZEICHEX)
                    For j = 20 To 2 Step -1
                        cFeld = SwapStr(cFeld, Space(j), " ")
                    Next j
                    rsrs!BEZEICH = Left(cFeld, 35)
                End If
            End If

            If Not IsNull(rsZ!LEKPREX) Then
                cLEKPR = rsZ!LEKPREX
                
                If Right(cLEKPR, 1) = "€" Then
                    rsrs!lekpr = Left(cLEKPR, Len(cLEKPR) - 1)
                Else
                    rsrs!lekpr = rsZ!LEKPREX
                End If
            End If
            
            'die optionalen abfragen
            
            For i = 0 To byAnzahlSpaltenEX - 1
            
                If sKissSpalteEX(i) = "MWSTEX" Then
                
                    If Not IsNull(rsZ!MWSTEX) Then
                    
                        cMWST = rsZ!MWSTEX
                        If Check5.Value = vbChecked Then
                            If cMWST = Trim(Text2(0).Text) Then
                                cMWST = "V"
                                rsrs!MWST = cMWST
                            ElseIf cMWST = Trim(Text2(1).Text) Then
                                cMWST = "E"
                                rsrs!MWST = cMWST
                            Else
                                rsrs!MWST = rsZ!MWSTEX
                            End If
                        Else
                            If cMWST = "19" Then
                                cMWST = "V"
                                rsrs!MWST = cMWST
                            ElseIf cMWST = "7" Then
                                cMWST = "E"
                                rsrs!MWST = cMWST
                            Else
                                rsrs!MWST = rsZ!MWSTEX
                            End If
        
                        End If
                    End If
                End If
                
                If sKissSpalteEX(i) = "MENGEEX" Then
                    If Not IsNull(rsZ!MENGEEX) Then
                        rsrs!Menge = Val(rsZ!MENGEEX)
                    End If
                End If
                
                If sKissSpalteEX(i) = "BESTANDEX" Then
                    If Not IsNull(rsZ!BESTANDEX) Then
                        rsrs!BESTAND = Val(rsZ!BESTANDEX)
                    End If
                End If
                
                If sKissSpalteEX(i) = "VPEEX" Then
                    If Not IsNull(rsZ!VPEEX) Then
                        rsrs!MINMEN = Val(rsZ!VPEEX)
                    End If
                End If
                
                If sKissSpalteEX(i) = "FARBNREX" Then
                    If Not IsNull(rsZ!FARBNREX) Then
                        rsrs!FARBNR = Val(rsZ!FARBNREX)
                    End If
                End If
                
                If sKissSpalteEX(i) = "INHALTEX" Then
                    If Not IsNull(rsZ!INHALTEX) Then
                        rsrs!INHALT = Val(rsZ!INHALTEX)
                    End If
                End If
                
                If sKissSpalteEX(i) = "INHALTSBEZEICHNUNGEX" Then
                    If Not IsNull(rsZ!INHALTSBEZEICHNUNGEX) Then
                        rsrs!INHALTBEZ = UCase(Left(Trim(CStr(rsZ!INHALTSBEZEICHNUNGEX)), 3))
                    End If
                End If
                
                If sKissSpalteEX(i) = "GROESSEEX" Then
                    If Not IsNull(rsZ!GROESSEEX) Then
                        rsrs!GROESSE = Left(Trim(CStr(rsZ!GROESSEEX)), 10)
                    End If
                End If
                
                If sKissSpalteEX(i) = "NOTIZENEX" Then
                    If Not IsNull(rsZ!NOTIZENEX) Then
                        cFeld = Trim(CStr(rsZ!NOTIZENEX))
                        For j = 20 To 2 Step -1
                            cFeld = SwapStr(cFeld, Space(j), " ")
                        Next j
                        rsrs!NOTIZEN = Left(cFeld, 40)
                    End If
                End If
                
                If sKissSpalteEX(i) = "LINIEEX" Then
                    If Not IsNull(rsZ!LINIEEX) Then
                    
                        If Val(rsZ!LINIEEX) < 500 Then
                            lLinie = Val(rsZ!LINIEEX) + 500
                        Else
                            lLinie = Val(rsZ!LINIEEX)
                        End If
                    
                        rsrs!LPZ = lLinie
                        
                        SchreibeLinieInLinbez lLinie, lLinr
                    End If
                End If
                
                If sKissSpalteEX(i) = "AGNEX" Then
                    If Not IsNull(rsZ!AGNEX) Then
                    
                        Dim lAgnInDatei As Long
                        lAgnInDatei = Val(rsZ!AGNEX)
                    
                        If Check6.Value = vbChecked Then
                            lBIOAGN = ermAgnoverBio(Val(rsZ!AGNEX))
                            If lBIOAGN > 0 Then
                                rsrs!AGN = lBIOAGN
                            Else
                                rsrs!AGN = lagn
                            End If
                        Else
                            If lAgnInDatei > 0 Then
                                rsrs!AGN = lAgnInDatei
                            Else
                                rsrs!AGN = lagn
                            End If
                        End If
                    End If
                
                End If
                
            Next i
           
            rsrs!linr = lLinr
            
            rsrs!RKZ = "N"
            rsrs!GEFUEHRT = "J"
            rsrs!GRUNDPREIS = "J"
            
            rsrs.Update
            
        rsZ.MoveNext
        Loop
    End If
    rsZ.Close
    rsrs.Close
    dbExcel.Close
    
    
    txtStatus.Text = 5
    anzeige "normal", "Die Datensätze werden überprüft...", Label1(4)
    
    
    '6. EAN - Duplikatsüberprüfung in der Importtabelle Anzahl ermitteln
    lAnz = CLng(ErmittlungMeisterDuplis)
    If lAnz > 0 Then
        anzeige "Rot", lAnz & " EAN - Duplikate wurden in der Importtabelle gefunden.", Label1(4)
    End If

    txtStatus.Text = 7
    '7. diverse Feldüberprüfungen vornehmen
    feldcheckMeister
    anzeige "normal", "Die Datensätze werden verarbeitet...", Label1(4)

    txtStatus.Text = 12
    '8. Tabelle IMPORTPRI zur Datenbank kopieren
    loeschNEW "Importpri", gdBase
    TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "IMPORTPRI"

    txtStatus.Text = 15

    EXCELStep3 = True
        
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 3349 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "EXCELStep3"
        Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        Resume Next
    End If
End Function
Private Sub SchreibeLinieInLinbez(lLpz As Long, lLinr As Long)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cLinienbez      As String

    If DatendrinSQL("Select * from linbez where lpz = " & lLpz & " and Linr = " & lLinr & " ", gdBase) Then
    
    Else
        cLinienbez = "Linie " & lLpz
        sSQL = "Insert into Linbez (LPZ,LINR,LINBEZEICH,Sorti) values  "
        sSQL = sSQL & " ( " & lLpz & " "
        sSQL = sSQL & " , " & lLinr & " "
        sSQL = sSQL & " , '" & cLinienbez & "' "
        sSQL = sSQL & " , " & lLpz & " "
        sSQL = sSQL & " ) "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "SchreibeLinieInLinbez"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub flex(krit As String)
On Error GoTo LOKAL_ERROR
    
    Dim lcount  As Long
    
    MSFlexGrid1.Redraw = False
    For lcount = 1 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Row = lcount

        Select Case krit
            Case "auswählen"
                MSFlexGrid1.Text = "ausgewählt"
                MSFlexGrid1.CellFontBold = True
                MSFlexGrid1.CellForeColor = vbGreen
                
            Case "entfernen"
                MSFlexGrid1.Text = "entfernt"
                MSFlexGrid1.CellFontBold = True
                MSFlexGrid1.CellForeColor = vbRed
        End Select
    Next lcount
    
    MSFlexGrid1.Redraw = True
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "flex"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub WerteGroßAnzeigen()
On Error GoTo LOKAL_ERROR

    Dim cSpaltenname    As String

    If List2.ListIndex < 0 Then
    
    Else
        Frame4.Visible = True
        Frame3.Visible = False
        
        Me.Refresh
    
        cSpaltenname = List2.list(List2.ListIndex)
        zeigWerte Check8.Value, Label2(5), Label2(8), cSpaltenname, List3, List4, "asc"
        anzeige "WARN", cSpaltenname, Label2(21)
    End If
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "WerteGroßAnzeigen"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Zuordnung(sSchemaname As String, iFremdnr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim sKissSpalte     As String
    Dim sKissDaba       As String
    Dim sFremdSpalte    As String
    Dim sFremdSpalteK   As String
    Dim lcount          As Long

    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    If Label3(0).Caption = "" Then
        anzeige "rot", "Bitte einen Eintrag auswählen!", Label1(4)
    Else
        sFremdSpalte = "[" & Label3(0).Caption & "]"
        sFremdSpalteK = Label3(0).Caption
        If Label3(1).Caption <> "" Then
            sFremdSpalte = sFremdSpalte & " & Space(1) & [" & Label3(1).Caption & "]"
            sFremdSpalteK = sFremdSpalteK & " + " & Label3(1).Caption
            If Label3(2).Caption <> "" Then
                sFremdSpalte = sFremdSpalte & " & Space(1) & [" & Label3(2).Caption & "]"
                sFremdSpalteK = sFremdSpalteK & " + " & Label3(2).Caption
            End If
        End If
    End If
    
    Label3(0).Caption = ""
    Label3(1).Caption = ""
    Label3(2).Caption = ""
    
    If List5.ListIndex < 0 Then
        anzeige "rot", "Bitte einen Eintrag auswählen!", Label1(4)
        Exit Sub
    Else
        For lcount = 0 To List5.ListCount - 1
            If List5.Selected(lcount) = True Then
            
                If InStr(1, List5.list(List5.ListIndex), " ") = 0 Then
                    sKissSpalte = List5.list(List5.ListIndex)
                Else
                    sKissSpalte = Trim(Mid$(List5.list(List5.ListIndex), 1, InStr(1, List5.list(List5.ListIndex), " ")))
                End If

                Select Case sKissSpalte
                    
                    Case "Lieferantenbestellnr"
                        sKissDaba = "LIBESNREX"
                    Case "EANCode(Strichcode)"
                        sKissDaba = "EANEX"
                    Case "Artikelbezeichnung"
                        sKissDaba = "BEZEICHEX"
                    Case "Listenverkaufspreis"
                        sKissDaba = "VKPREX"
                    Case "Listeneinkaufspreis"
                        sKissDaba = "LEKPREX"
                    Case "MwSt.-Kennzeichen"
                        sKissDaba = "MWSTEX"
                    Case "Notizen"
                        sKissDaba = "NOTIZENEX"
                    Case "Größe"
                        sKissDaba = "GROESSEEX"
                    Case "Farbnr"
                        sKissDaba = "FARBNREX"
                    Case "VPE"
                        sKissDaba = "VPEEX"
                    Case "Linie"
                        sKissDaba = "LINIEEX"
                    Case "AGN"
                        sKissDaba = "AGNEX"
                    Case "Inhalt"
                        sKissDaba = "INHALTEX"
                    Case "Inhaltsbezeichnung"
                        sKissDaba = "INHALTSBEZEICHNUNGEX"
                    Case "Menge"
                        sKissDaba = "MENGEEX"
                    Case "Bestand"
                        sKissDaba = "BESTANDEX"
                        
                        
                End Select
                    
            End If
        Next lcount
    End If
    
    If sKissSpalte <> "" And sFremdSpalte <> "" Then
    
        sSQL = "Delete from SPZUORD where ZUKISSSP = '" & sKissSpalte & "'"
        sSQL = sSQL & " and Schemaname = '" & sSchemaname & "'"
        gdBase.Execute sSQL, dbFailOnError
    
        sSQL = "Insert into SPZUORD (ZUKISSSP,ZUKISSDABA, ZUFREMDSP,FREMDKLAR,SCHEMANAME,FREMDNR"
        sSQL = sSQL & ",LASTDATE,LASTTIME"
        sSQL = sSQL & " ) "
        sSQL = sSQL & " values "
        sSQL = sSQL & "('" & sKissSpalte & "'"
        sSQL = sSQL & ",'" & sKissDaba & "'"
        sSQL = sSQL & ",'" & sFremdSpalte & "'"
        sSQL = sSQL & ",'" & sFremdSpalteK & "'"
        sSQL = sSQL & ",'" & sSchemaname & "'"
        sSQL = sSQL & ", " & iFremdnr & " "
        sSQL = sSQL & ",'" & DateValue(Now) & "'"
        sSQL = sSQL & ",'" & TimeValue(Now) & "'"
        sSQL = sSQL & ")"
        gdBase.Execute sSQL, dbFailOnError
        
        List5.RemoveItem List5.ListIndex
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zuordnung"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Speichervorauswahl(checkX As CheckBox, sSchemaname As String, labelx As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    sSQL = "Delete from SPZUORDVOR where SCHEMANAME = '" & sSchemaname & "'"
    gdBase.Execute sSQL, dbFailOnError

    If checkX.Value = vbChecked Then
        sSQL = "Insert into SPZUORDVOR (ZUFREMDSP,SCHEMANAME) "
        sSQL = sSQL & " values "
        sSQL = sSQL & "('" & checkX.Caption & "'"
        sSQL = sSQL & ",'" & sSchemaname & "'"
        sSQL = sSQL & ")"
        gdBase.Execute sSQL, dbFailOnError
        anzeige "warn", "", labelx
    Else
        anzeige "warn", "Möchten Sie bestimmte Artikelbereiche auswählen, dann markieren(klicken) Sie in der linken Spaltenauswahl eine Spalte.", labelx
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Speichervorauswahl"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function checkZuordnung(sSchemaname As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSpalten(5) As String
    Dim cVorauswahlspalte As String
    checkZuordnung = False
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    sSpalten(0) = "Lieferantenbestellnr"
    sSpalten(1) = "EANCode(Strichcode)"
    sSpalten(2) = "Artikelbezeichnung"
    sSpalten(3) = "Listenverkaufspreis"
    sSpalten(4) = "Listeneinkaufspreis"

    For i = 0 To 4
        If checkKissspalte(sSchemaname, sSpalten(i)) Then
            checkZuordnung = True
        Else
            checkZuordnung = False
            Exit For
        End If
    Next i
    
    'check auch die Vorauswahlspalte
    cVorauswahlspalte = ermVorauswahl(Label2(5).Caption & " " & Label2(8).Caption)
        
    If cVorauswahlspalte = "" Then
        If checkZuordnung = True Then
            anzeige "rot2", "Jetzt noch die Vorauswahlspalte benennen!", Label1(4)
            Check2.BackColor = vbRed
            checkZuordnung = False
        End If
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkZuordnung"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function checkKissspalte(sSchemaname As String, sSpalte As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    checkKissspalte = False
    
    sSQL = "Select * from SPZUORD where ZUKISSSP = '" & sSpalte & "'"
    sSQL = sSQL & " and SCHEMANAME = '" & sSchemaname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        checkKissspalte = True
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "checkKissspalte"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub füllelist5()
    On Error GoTo LOKAL_ERROR

    List5.Clear
    List5.AddItem "Lieferantenbestellnr Pflichtangabe"
    List5.AddItem "EANCode(Strichcode)  Pflichtangabe"
    List5.AddItem "Artikelbezeichnung   Pflichtangabe"
    List5.AddItem "Listenverkaufspreis  Pflichtangabe"
    List5.AddItem "Listeneinkaufspreis  Pflichtangabe"
    List5.AddItem "MwSt.-Kennzeichen"
    List5.AddItem "Notizen"
    List5.AddItem "Größe"
    List5.AddItem "Farbnr"
    List5.AddItem "VPE"
    List5.AddItem "Linie"
    List5.AddItem "AGN"
    List5.AddItem "Inhalt"
    List5.AddItem "Inhaltsbezeichnung"
    List5.AddItem "Menge -> für Bestelldatei"
    List5.AddItem "Bestand"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "füllelist5"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub LoeschInlist5(Listx As ListBox)
    On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim j As Integer
    Dim sSpalte As String
    
    If Listx.ListCount - 1 > 0 Then
        For i = 0 To Listx.ListCount - 1
            sSpalte = Trim(Left(Listx.list(i), 25))
            For j = 0 To List5.ListCount - 1
                If UCase(sSpalte) = UCase(Trim(List5.list(j))) Then
                    List5.RemoveItem j
                End If
            Next j
        Next i
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "LoeschInlist5"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherpfadEXCEL(spfa As String)
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    
    sSQL = "Delete from  EXCELe "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into EXCELe (EXCELpfad) values ('" & spfa & " ' ) "
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherpfadEXCEL"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub pfadseekExcel()
    On Error GoTo LOKAL_ERROR
    
    Dim sTitle      As String
    Dim sFilter     As String
    Dim sOldpfad    As String

    sTitle = "Speichern des Pfades"
    
    sFilter = "Excel - Dateien (*.xls)|*.xls"
    
    sOldpfad = Label2(1).Caption
    sExcelpfad = pfadaendern(sTitle, sFilter, sOldpfad)
    
    Label2(1).Caption = sExcelpfad
    
    sExcelpfad = ShortPath(sExcelpfad)
    
    speicherpfadEXCEL Trim(sExcelpfad)
    filefreshEXCEL
    zeigeanzahl File1
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "pfadseekExcel"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub filefreshEXCEL()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    
    cPfad = gcDBPfad        'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    ermittleEXCeLpfad
    
    File1.Path = sExcelpfad
    File1.Pattern = "*.xls"
    File1.Refresh

    Exit Sub
LOKAL_ERROR:
    If err.Number = 68 Or err.Number = 76 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "filefreshEXCEL"
        Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub ermittleEXCeLpfad()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    'TchPfad ermitteln
    sExcelpfad = ""
    Set rsrs = gdBase.OpenRecordset("EXCELE", dbOpenTable)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Excelpfad) Then
            sExcelpfad = Trim(rsrs!Excelpfad)
        End If
    End If
    rsrs.Close
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittleEXCeLpfad"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub EXCELStep1()
    On Error GoTo LOKAL_ERROR
    
    If Not NewTableSuchenDBKombi("EXCELE", gdBase) Then 'das erste Mal
        CreateTableT2 "EXCELE", gdBase
        speicherpfadEXCEL gsKinPfad
        speicherGrundETransfEXCEL
    End If
    
    If Not NewTableSuchenDBKombi("SPZUORD", gdBase) Then 'das erste Mal
        CreateTableT2 "SPZUORD", gdBase
    End If
    
    If Not NewTableSuchenDBKombi("SPZUORDLINR", gdBase) Then 'das erste Mal
        CreateTableT2 "SPZUORDLINR", gdBase
    End If
    
    If Not NewTableSuchenDBKombi("SPZUORDVOR", gdBase) Then 'das erste Mal
        CreateTableT2 "SPZUORDVOR", gdBase
    End If
    
    If Not NewTableSuchenDBKombi("SPZUORDDETAILS", gdBase) Then 'das erste Mal
        CreateTableT2 "SPZUORDDETAILS", gdBase
    End If
    
    ermittleEXCeLpfad
    
    If sExcelpfad = "" Then
        anzeige "Normal", "Keine Stammdaten - Dateien vorhanden.", Label1(4)
    Else
        Label2(1).Caption = sExcelpfad
        Label2(1).Refresh
        
        filefreshEXCEL
        zeigeanzahl File1
        Frame1.Visible = True
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "EXCELStep1"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub speicherGrundETransfEXCEL()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim i As Integer
    
    For i = 0 To 9
        sSQL = "Update EXCELE set transf" & i & "= true "
        gdBase.Execute sSQL, dbFailOnError
    Next i
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "speicherGrundETransfEXCEL"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigeanzahl(fil As FileListBox)
    On Error GoTo LOKAL_ERROR
    
    If File1.ListCount = 0 Then
        anzeige "Rot2", "Keine Stammdaten - Dateien vorhanden.", Label1(4)
    ElseIf File1.ListCount = 1 Then
'        anzeige "Normal", File1.ListCount & " Stammdaten - Datei vorhanden.", Label1(4)
        anzeige "Normal", File1.ListCount & " Datei vorhanden.", Label2(13)
    Else
        anzeige "Normal", File1.ListCount & " Stammdaten - Dateien vorhanden.", Label1(4)
        anzeige "Normal", File1.ListCount & " Dateien vorhanden.", Label2(13)
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "zeigeanzahl"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigTabellenblätter(cDatname As String, Listx As ListBox)
On Error GoTo LOKAL_ERROR

    Listx.Clear
    
    Dim gsExcel50   As String
    Dim dbExcel     As Database
    Dim lcount   As Long
    Dim cTabname    As String
    
    Command5(6).Enabled = False
    Command5(6).BackColor = Command5(0).BackColor
    
    anzeige "normal", "", Label1(4)
        
    gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
        
    Set dbExcel = OpenDatabase(sExcelpfad & "\" & cDatname, 0, 0, gsExcel50)
    
    For lcount = 0 To dbExcel.TableDefs.Count - 1
        cTabname = dbExcel.TableDefs(lcount).name
        Listx.AddItem cTabname
    Next lcount
    
    dbExcel.Close
    Command5(6).Enabled = True
    Command5(6).BackColor = vbRed
    
    'Anzahl Tabblätter zeigen
    If Listx.ListCount = 0 Then
        anzeige "Normal", "kein Tabellenblatt vorhanden.", Label2(14)
    ElseIf Listx.ListCount = 1 Then
        anzeige "Normal", Listx.ListCount & " Tabellenblatt vorhanden.", Label2(14)
    Else
        anzeige "Normal", Listx.ListCount & " Tabellenblätter vorhanden.", Label2(14)
    End If
       
Exit Sub
LOKAL_ERROR:
    If err.Number = 3274 Or err.Number = 3170 Then
    
        Select Case gsExcel50
            Case Is = "Excel 3.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 4.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 4.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 5.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 7.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 7.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 8.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 8.0;HDR=yes;IMEX=1;"
                'Geht nicht, noch unbekannt warum
                anzeige "rot1", "Diese Datei hat nicht das erwartete Format.", Label1(4)
        End Select
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "zeigTabellenblätter"
        Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub zeigSpalten(cDatname As String, cTabBlattname As String, Listx As ListBox)
On Error GoTo LOKAL_ERROR

    Listx.Clear
    
    Dim gsExcel50   As String
    Dim dbExcel     As Database
    Dim tdTd        As TableDef
    Dim lcount      As Long
    Dim cTabname    As String
    Dim bfind       As Boolean
    Dim cFeldName   As String
        
    Command5(5).Enabled = False
    Command5(5).BackColor = Command5(0).BackColor
    
    anzeige "normal", "", Label1(4)
    
    bfind = False
    gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
        
    Set dbExcel = OpenDatabase(sExcelpfad & "\" & cDatname, 0, 0, gsExcel50)
    
    Set tdTd = dbExcel.TableDefs(cTabBlattname)
    For lcount = 0 To dbExcel.TableDefs.Count - 1
        cTabname = dbExcel.TableDefs(lcount).name
        
        
        If UCase$(cTabname) = UCase(cTabBlattname) Then
            bfind = True
            Exit For
        Else
            bfind = False
        End If
        
    Next lcount
    
    If bfind Then
        For lcount = 0 To tdTd.Fields.Count - 1
            cFeldName = tdTd.Fields(lcount).name
            
            If DatendrinSQL("select [" & cFeldName & "] from [" & cTabname & "] where not [" & cFeldName & "] is null", dbExcel) Then
                Listx.AddItem cFeldName
            Else
                
'                listx.AddItem cFeldName & " (leer)"
            End If
        Next lcount
    
    End If
    
    dbExcel.Close
        
    Command5(5).Enabled = True
    Command5(5).BackColor = vbRed
    
    'Anzahl Spalten zeigen
    If Listx.ListCount = 0 Then
        anzeige "Normal", "keine Spalte vorhanden.", Label2(15)
    ElseIf Listx.ListCount = 1 Then
        anzeige "Normal", Listx.ListCount & " Spalte vorhanden.", Label2(15)
    Else
        anzeige "Normal", Listx.ListCount & " Spalten vorhanden.", Label2(15)
    End If
    
       
Exit Sub
LOKAL_ERROR:
    If err.Number = 3274 Or err.Number = 3170 Then
    
        Select Case gsExcel50
            Case Is = "Excel 3.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 4.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 4.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 5.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 7.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 7.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 8.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 8.0;HDR=yes;IMEX=1;"
                'Geht nicht noch unbekannt warum
                anzeige "rot1", "Diese Datei hat nicht das erwartete Format.", Label1(4)
        End Select
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "zeigSpalten"
        Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub zeigWerte(bEindeut As Boolean, cDatname As String, cTabBlattname As String, cSpalte As String, Listx As ListBox, listxUE As ListBox, sOrder As String)
On Error GoTo LOKAL_ERROR

    Listx.Clear
    
    Dim gsExcel50   As String
    Dim dbExcel     As Database
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lcount      As Long
    Dim cSatz       As String
    Dim cFeld       As String
    Dim lZaehler     As Long
    
'    Command5(9).Enabled = False
'    Command5(9).BackColor = Command5(0).BackColor
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 11
    
    gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
        
    Set dbExcel = OpenDatabase(sExcelpfad & "\" & cDatname, 0, 0, gsExcel50)

    If bEindeut Then
        sSQL = " Select  [" & cSpalte & "],count([" & cSpalte & "]) as count from [" & cTabBlattname & "]"
        sSQL = sSQL & " where not [" & cSpalte & "] is null "
        sSQL = sSQL & " group by [" & cSpalte & "] "
    Else
        sSQL = " Select [" & cSpalte & "] from [" & cTabBlattname & "] where not [" & cSpalte & "] is null "
    End If
    
    Listx.Clear
    listxUE.Clear
    
    listxUE.Visible = False
    Listx.Visible = False
    
    If bEindeut Then
        listxUE.AddItem "Wertangaben" & Space(24) & "Anzahl Artikel"
    Else
        listxUE.AddItem "Wertangaben"
    End If
    
    lZaehler = 0
    Set rsrs = dbExcel.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cSatz = ""
            If Not IsNull(rsrs.Fields(0)) Then
                cFeld = Trim(rsrs.Fields(0))
                
                If Len(cFeld) > 34 Then
                    cFeld = Left(cFeld, 31) & "..."
                End If
            
                cSatz = cFeld & Space(35 - Len(cFeld))
                If bEindeut Then
                    If Not IsNull(rsrs.Fields(1)) Then
                        cFeld = Trim(rsrs.Fields(1))
                    Else
                        cFeld = ""
                    End If
                    cSatz = cSatz & cFeld & " Artikel"
                End If
                
                Listx.AddItem cSatz
                lZaehler = lZaehler + 1
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    dbExcel.Close
    
    If bEindeut Then
        anzeige "normal", lZaehler & " eindeutige Datensätze", Label2(4)
    Else
        anzeige "normal", lZaehler & " Datensätze", Label2(4)
    End If
    
    Listx.Visible = True
    listxUE.Visible = True
        
    Screen.MousePointer = 0
       
Exit Sub
LOKAL_ERROR:
    If err.Number = 3274 Or err.Number = 3170 Then
    
        Select Case gsExcel50
            Case Is = "Excel 3.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 4.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 4.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 5.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 7.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 7.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 8.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 8.0;HDR=yes;IMEX=1;"
                'Geht nicht noch unbekannt warum
                anzeige "rot1", "Diese Datei hat nicht das erwartete Format.", Label1(4)
        End Select
    ElseIf err.Number = 3349 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "zeigWerte"
        Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub Zuordnunganzeigen(Listx As ListBox, sSchemaname As String)
On Error GoTo LOKAL_ERROR

    Listx.Clear
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim cSatz       As String
    Dim cFeld       As String
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 11
    
    sSQL = "Select * from SPZUORD where SCHEMANAME = '" & sSchemaname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cSatz = ""
            If Not IsNull(rsrs!ZUKISSSP) Then
                cFeld = rsrs!ZUKISSSP
            Else
                cFeld = ""
            End If
            cSatz = cFeld & Space(35 - Len(cFeld)) & "<--> "
            
            If Not IsNull(rsrs!FREMDKLAR) Then
                cFeld = rsrs!FREMDKLAR
            Else
                cFeld = ""
            End If
            If Len(cFeld) > 35 Then
                cFeld = Left(cFeld, 32) & "..."
            End If
            
            cSatz = cSatz & cFeld & Space(35 - Len(cFeld))
            
            If Not IsNull(rsrs!FREMDNR) Then
                cFeld = rsrs!FREMDNR
            Else
                cFeld = "0"
            End If
            
            If Val(cFeld) > 0 Then
                cFeld = "Besonderheiten"
            Else
                cFeld = ""
            End If
            cSatz = cSatz & cFeld & Space(15 - Len(cFeld))
            
            Listx.AddItem cSatz
            
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
        
    Screen.MousePointer = 0
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zuordnunganzeigen"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Vorauswahlcheck(sSchemaname As String, checkX As CheckBox, labelx As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim cFeld       As String
    
    anzeige "normal", "", Label1(4)
    Screen.MousePointer = 11
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    checkX.Caption = ""
    cFeld = "Möchten Sie bestimmte Artikelbereiche auswählen, dann markieren(klicken) Sie in der linken Spaltenauswahl eine Spalte."
    
    sSQL = "Select * from SPZUORDVOR where SCHEMANAME = '" & sSchemaname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!ZUFREMDSP) Then
            checkX.Caption = rsrs!ZUFREMDSP
            checkX.Value = vbChecked
            cFeld = ""
        End If
    End If
    rsrs.Close
    
    anzeige "warn", cFeld, labelx

    Screen.MousePointer = 0
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Vorauswahlcheck"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function ermVorauswahl(sSchemaname As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    Screen.MousePointer = 11
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    ermVorauswahl = ""

    sSQL = "Select * from SPZUORDVOR where SCHEMANAME = '" & sSchemaname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!ZUFREMDSP) Then
            ermVorauswahl = rsrs!ZUFREMDSP
        End If
    End If
    rsrs.Close
    
    Screen.MousePointer = 0
       
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermVorauswahl"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub Zuordnungloeschen(sSchemaname As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    
    anzeige "normal", "", Label1(4)
    
    sSchemaname = SwapStr(sSchemaname, "'", "")
    
    sSQL = "Delete from SPZUORD where SCHEMANAME = '" & sSchemaname & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from SPZUORDVOR where SCHEMANAME = '" & sSchemaname & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from SPZUORDLINR where SCHEMANAME = '" & sSchemaname & "'"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Zuordnungloeschen"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub ermittlespalten()
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer

    For i = 0 To byAnzahlSpalten
        Select Case UCase(sSpaltenbez(i))
            Case Is = "ARTNR"
            SpaltennummerArtnr = i
            Case Is = "BEZEICH"
            SpaltennummerBEZEICH = i
            Case Is = "LINR"
            SpaltennummerLINR = i
            Case Is = "LPZ"
            SpaltennummerLPZ = i
            Case Is = "LIBESNR"
            SpaltennummerLIBESNR = i
            Case Is = "LEKPR"
            SpaltennummerLEKPR = i
            Case Is = "VKPR"
            SpaltennummerVKPR = i
            Case Is = "KVKPR1"
            SpaltennummerKVKPR1 = i
            Case Is = "MINBEST"
            SpaltennummerMINBEST = i
            Case Is = "GEFUEHRT"
            SpaltennummerGEFUEHRT = i
            Case Is = "RABATT_OK"
            SpaltennummerRABATT_OK = i
            Case Is = "PREISSCHU"
            SpaltennummerPREISSCHU = i
            Case Is = "NOTIZEN"
            SpaltennummerNOTIZEN = i
            Case Is = "AGN"
            SpaltennummerAGN = i
            Case Is = "PGN"
            SpaltennummerPGN = i
            Case Is = "RKZ"
            SpaltennummerRKZ = i
            Case Is = "EAN"
            SpaltennummerEAN = i
            Case Is = "MINMEN"
            SpaltennummerMINMEN = i
            Case Is = "BESTAND"
                SpaltennummerBESTAND = i
            Case Is = "MENGE"
                SpaltennummerMENGE = i
            Case Is = "MWST"
            SpaltennummerMWST = i
            Case Is = "MNOTIZEN"
            SpaltennummerMNOTIZEN = i
            Case Is = "GROESSE"
            SpaltennummerGROESSE = i
        End Select
        Select Case UCase(sSpaltenname(i))
            Case Is = "KVK NEU"
            SpaltennummerKVKNEU = i
        End Select
    Next i
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermittlespalten"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
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
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "Tabellenbreiteanpassen"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function FormatiereBildschirmdaten() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim lreservArtnr    As Long
    Dim lvergebeArtnr   As Long

    Dim cSQL            As String

    Dim dNettospanne    As Double
    Dim dEK             As Double
    Dim cMWST           As String
    Dim cNewKassenPr    As String
    Dim siAnzeige       As Single
    Dim lAnz            As Long
    
    
    txtStatus.Text = 18
    
    FormatiereBildschirmdaten = False
    anzeige "normal", "Neue Artikel werden ermittelt...", Label1(4)
    'Farbe alle auf neu
    sSQL = "Update ImportPri set AWM = '98' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 21
    
    sSQL = "Create Index AWM on IMPORTPRI (AWM)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index ean on IMPORTPRI (ean)"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 23
    
    'Artikel mit EAN übereinstimmung auf standard
    sSQL = "Update ImportPri inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN = ImportPri.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " ImportPri.AWM = '0' "
    sSQL = sSQL & " ,ImportPri.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,ImportPri.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,ImportPri.Bestand = Artikel.Bestand  "
'    sSQL = sSQL & " ,ImportPri.AGN = Artikel.AGN  "
    sSQL = sSQL & " ,ImportPri.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,ImportPri.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where Importpri.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 25
    
    sSQL = "Update ImportPri inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN2 = ImportPri.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRI.AWM = '0' "
    sSQL = sSQL & " ,ImportPri.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,ImportPri.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,ImportPri.Bestand = Artikel.Bestand  "
'    sSQL = sSQL & " ,ImportPri.AGN = Artikel.AGN  "
    sSQL = sSQL & " ,ImportPri.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,ImportPri.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where Importpri.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 28
    
    sSQL = "Update ImportPri inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.EAN3 = ImportPri.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRI.AWM = '0' "
    sSQL = sSQL & " ,ImportPri.KVKALT = Artikel.kvkpr1  "
    sSQL = sSQL & " ,ImportPri.artnr = Artikel.artnr  "
    sSQL = sSQL & " ,ImportPri.Bestand = Artikel.Bestand  "
'    sSQL = sSQL & " ,ImportPri.AGN = Artikel.AGN  "
    sSQL = sSQL & " ,ImportPri.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,ImportPri.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where Importpri.ean  <> '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    sSQL = "Update ImportPri inner join ARTEAN_K on "
    sSQL = sSQL & " ARTEAN_K.EAN = ImportPri.EAN "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " IMPORTPRI.AWM = '0' "
    sSQL = sSQL & " ,ImportPri.artnr = ARTEAN_K.artnr  "
    sSQL = sSQL & " where Importpri.ean  <> '0' "
    sSQL = sSQL & " and Importpri.artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update ImportPri inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = ImportPri.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " ImportPri.KVKALT = Artikel.kvkpr1 "
    sSQL = sSQL & " ,ImportPri.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,ImportPri.Rabatt_ok = Artikel.Rabatt_ok  "
    sSQL = sSQL & " ,ImportPri.gefuehrt = Artikel.gefuehrt  "
    sSQL = sSQL & " where Importpri.ean  <> '0' and ImportPri.AWM = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    
    
    'where importpri.ean = '0' der macht ein check in artlief
    'auf linr und libesnr stimmigkeit 600283
    
    sSQL = "Update ImportPri inner join ARTLIEF on "
    sSQL = sSQL & " ARTLIEF.LINR = ImportPri.LINR and Trim(ARTLIEF.LIBESNR) = Trim(ImportPri.LIBESNR) "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " ImportPri.AWM = '0' "
    sSQL = sSQL & " ,ImportPri.artnr = ARTLIEF.artnr  "
    sSQL = sSQL & " where Importpri.ean  = '0' "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 29

    sSQL = "Update ImportPri inner join Artikel on "
    sSQL = sSQL & " ARTIKEL.ARTNR = ImportPri.ARTNR "
    sSQL = sSQL & "Set "
    sSQL = sSQL & " ImportPri.KVKALT = Artikel.kvkpr1  "

    sSQL = sSQL & " ,ImportPri.Bestand = Artikel.Bestand  "
    sSQL = sSQL & " ,ImportPri.gefuehrt = Artikel.gefuehrt  "

    sSQL = sSQL & " where Importpri.ean  = '0' and ImportPri.AWM = '0' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 32
    
    sSQL = "Update ImportPri set AWM = '96' "
    sSQL = sSQL & " where LPZ = 70 and ETIMERK = 'J' and not awm = '98'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 33
    
    sSQL = "Update ImportPri set AWM = '95' "
    sSQL = sSQL & " where LPZ = 70 and ETIMERK = 'J' and awm = '98'"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 35
    
    If NewTableSuchenDBKombi("importdupli", gdBase) Then
    
        sSQL = "Update ImportPri inner join importdupli on "
        sSQL = sSQL & " importdupli.ean = ImportPri.ean set AWM = '94' "
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    
    cSQL = " Update ImportPri inner join Artikel on "
    cSQL = cSQL & " ARTIKEL.ARTNR = ImportPri.ARTNR "
    cSQL = cSQL & "Set "
    cSQL = cSQL & " ImportPri.PREISSCHU = ARTIKEL.PREISSCHU "
    gdBase.Execute cSQL, dbFailOnError
    
    txtStatus.Text = 36
    '*****'
    'ab Hier check der Autokalkulierung
    '1. Unter Voreinstellung basierend auf LEK
    If gsSpanne = "LEK" Then
        '2. Etimerk = J
        '3. Not Preisschu
        '4. Feld Nettospanne gefüllt
        '5. LEK gefüllt
        
        cSQL = " Update ImportPri inner join Artikel on "
        cSQL = cSQL & " ARTIKEL.ARTNR = ImportPri.ARTNR "
        cSQL = cSQL & "Set "
        cSQL = cSQL & " ImportPri.PREISSCHU = ARTIKEL.PREISSCHU "
        cSQL = cSQL & " where ARTIKEL.PREISSCHU = 'N'"
        cSQL = cSQL & " and ImportPri.LEKPR > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        txtStatus.Text = 37
    
        
        cSQL = " Update ImportPri inner join ARTLIEF on "
        cSQL = cSQL & " ARTLIEF.ARTNR = ImportPri.ARTNR and ARTLIEF.LINR = ImportPri.LINR"
        cSQL = cSQL & " Set "
        cSQL = cSQL & " ImportPri.SPANNE = ARTLIEF.SPANNE "
        gdBase.Execute cSQL, dbFailOnError
        
        txtStatus.Text = 39
        
        cSQL = " Update ImportPri "
        cSQL = cSQL & " Set ETIMERK = 'J' "
        cSQL = cSQL & " where ImportPri.SPANNE > 0 "
        gdBase.Execute cSQL, dbFailOnError
        
        txtStatus.Text = 40
        
        cSQL = "Select * from ImportPri where ETIMERK = 'J'and (not ImportPri.Spanne is null or ImportPri.Spanne <> 0 ) "
        Set rsrs = gdBase.OpenRecordset(cSQL)

        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
                'Wir brauchen den LEK aus Mastemp
                'MWST
                'Nettospanne
                If Not IsNull(rsrs!lekpr) Then
                    dEK = rsrs!lekpr
                Else
                    dEK = 0
                End If

                If Not IsNull(rsrs!MWST) Then
                    cMWST = rsrs!MWST
                Else
                    cMWST = "V"
                End If


                If Not IsNull(rsrs!SPANNE) Then
                    dNettospanne = rsrs!SPANNE
                Else
                    dNettospanne = 0
                End If
                cNewKassenPr = Runden(CDbl(fnVKneuNS(dEK, cMWST, dNettospanne)))

                rsrs.Edit
                rsrs!KVKNEU = cNewKassenPr
                rsrs.Update

                rsrs.MoveNext
            Loop
        End If
        rsrs.Close
        
        txtStatus.Text = 46
        
        cSQL = " Update ImportPri "
        cSQL = cSQL & "Set  ImportPri.AWM = '97' "
        cSQL = cSQL & " where ETIMERK = 'J' and PREISSCHU = 'N'"
        gdBase.Execute cSQL, dbFailOnError
        
        txtStatus.Text = 50
    End If
    
    Dim dEkpr As Double
    
    txtStatus.Text = 53
    anzeige "normal", "Für neue Artikel werden freie Artikelnummern ermittelt...", Label1(4)
    
    
    lreservArtnr = HoleFreieArtikelNrab(glartv, glartb)
    
    txtStatus.Text = 0
    
    sSQL = "Select * from ImportPri where awm = '98' or awm = '95' or awm = '94'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnz = rsrs.RecordCount
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!artnr = lreservArtnr
            rsrs.Update
            
            lvergebeArtnr = NextfreieArtnr(lreservArtnr, glartb)
            If lvergebeArtnr = 0 Then
                anzeige "rot", "Es stehen keine neuen Artikelnummern zur Verfügung (Einstellungen überprüfen).", Label1(4)
                Exit Function
            Else
                lreservArtnr = lvergebeArtnr
                
                siAnzeige = siAnzeige + 1
                txtStatus.Text = CStr((100 * siAnzeige) / lAnz)
                
            End If
            

            rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close
    
    'dupliprüfung
    
    If NewTableSuchenDBKombi("importdupli", gdBase) Then
    
        sSQL = "Select ean from Importdupli"
        
        Dim cEANa As String
        lreservArtnr = HoleFreieArtikelNrab(glartv, glartb)
    
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                If Not IsNull(rsrs!EAN) Then
                    cEANa = rsrs!EAN
                End If
            
                sSQL = "Update ImportPri set artnr = " & lreservArtnr
                sSQL = sSQL & " where ean = '" & cEANa & "'"
                gdBase.Execute sSQL, dbFailOnError
            
                lvergebeArtnr = NextfreieArtnr(lreservArtnr, glartb)
                If lvergebeArtnr = 0 Then
                    
                    anzeige "rot", "Es stehen keine neuen Artikelnummern zur Verfügung (Einstellungen überprüfen).", Label1(4)
                    Exit Function
                Else
                    lreservArtnr = lvergebeArtnr
                End If
                rsrs.MoveNext
            Loop
        
        End If
        rsrs.Close
    
    End If
    
    FormatiereBildschirmdaten = True

    Exit Function
LOKAL_ERROR:
    If err.Number = 3372 Or err.Number = 53 Or err.Number = 3376 Or err.Number = 3375 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "FormatiereBildschirmdaten"
        Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
        Fehlermeldung1
        Resume Next
    End If
End Function
Private Sub Form_Load()
On Error GoTo LOKAL_ERROR

    positionieren159
    
    Skalieren Me, True, True: Schrift Me:
    Farbform Me, lblUeberschrift
    LogtoStart Me
    
    EXCELStep1
    
    lesenEinstellungen
    
    
    
    Label3(0).ForeColor = vbRed
    Label3(1).ForeColor = vbRed
    Label3(2).ForeColor = vbRed
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "Command5_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub lesenEinstellungen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    glartv = 600000
    glartb = 700000
    
    If NewTableSuchenDBKombi("FFE", gdBase) = True Then
        Set rsrs = gdBase.OpenRecordset("FFE", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            
            If Not IsNull(rsrs!ARTNRV) Then
                glartv = rsrs!ARTNRV
            Else
                glartv = 600000
            End If
            
            If Not IsNull(rsrs!ARTNRB) Then
                glartb = rsrs!ARTNRB
            Else
                glartb = 700000
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "lesenEinstellungen"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub


Private Sub PRIFlex_DblClick()
On Error GoTo LOKAL_ERROR
    
    sortierenHGrid PRIFlex
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PRIFlex_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
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
    Fehler.gsFunktion = "txtStatus_Change"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub positionieren159()
On Error GoTo LOKAL_ERROR

    posiFramex Frame1
    posiFramex Frame2
    posiFramex Frame3
    posiFramex Frame4
    posiFramex Frame6
    posiFramex Frame7
    
    posiFramex Frame8
    
    posiFramex Frame9
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "positionieren159"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub posiFramex(FrameX As Frame)
On Error GoTo LOKAL_ERROR

    FrameX.Top = 960
    FrameX.Left = 120
    FrameX.Height = 6135
    FrameX.Width = 11655
    FrameX.Caption = ""
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "posiFramex"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub feldcheckMeister()
    On Error GoTo LOKAL_ERROR
    
    Dim rsMeister       As Recordset
    Dim rsIMPORTPRI     As Recordset
    Dim sSQL            As String
    Dim sWeek           As String
    Dim iWeek           As Integer
    Dim sBez            As String
    Dim sEAN            As String
    
    loeschNEW "IMPORTPRI", gdApp
    CreateTable "IMPORTPRI", gdApp
    
    SpalteAnfuegenNEW "IMPORTPRI", "MENGE", "LONG", gdApp
    
    Set rsIMPORTPRI = gdApp.OpenRecordset("IMPORTPRI")
    
    Set rsMeister = gdApp.OpenRecordset("Meister")
    If Not rsMeister.EOF Then
        rsMeister.MoveFirst
        Do While Not rsMeister.EOF
            rsIMPORTPRI.AddNew
            rsIMPORTPRI!artnr = rsMeister!artnr
            rsIMPORTPRI!Menge = rsMeister!Menge
            
            'Bezeichnung auf * und ' checken
            If Not IsNull(rsMeister!BEZEICH) Then
                sBez = Left(rsMeister!BEZEICH, 35)
                sBez = SwapStr(sBez, "*", " ")  'stern
                sBez = SwapStr(sBez, Chr(34), "Z") '"
                sBez = SwapStr(sBez, "'", " ")  'Hochkommata
                sBez = SwapStr(sBez, ",", ".")  'komma
                sBez = SwapStr(sBez, "á", "ß")  'ß
                sBez = SwapStr(sBez, "”", "ö")  '
                sBez = SwapStr(sBez, "„", "ä")  '
                sBez = SwapStr(sBez, "", "ü")
                sBez = SwapStr(sBez, "™", "Ö")
                sBez = SwapStr(sBez, "š", "Ü")
                sBez = SwapStr(sBez, "`", " ")  '
                sBez = SwapStr(sBez, "|", " ")  '
                
                sBez = SwapStr(sBez, "Ã¤", "ä")  '
                sBez = SwapStr(sBez, "ÃŸ", "ß")  'ß
                sBez = SwapStr(sBez, "Ã¼", "ü")
                sBez = SwapStr(sBez, "Ã¶", "ö")  '
                sBez = SwapStr(sBez, "Ã–", "Ö")  '
            Else
                sBez = ""
            End If
            
            If gbTagAkt = True Then
                rsIMPORTPRI!BEZEICH = UCase(sBez)
            Else
                rsIMPORTPRI!BEZEICH = sBez
            End If
            
            
            
            If Not IsNull(rsMeister!LPZ) Then
                rsIMPORTPRI!LPZ = rsMeister!LPZ
            Else
                rsIMPORTPRI!LPZ = 1
            End If
            
            
            If Not IsNull(rsMeister!AGN) Then
                rsIMPORTPRI!AGN = rsMeister!AGN
            Else
                rsIMPORTPRI!AGN = 0
            End If
            
            If Not IsNull(rsMeister!PGN) Then
                rsIMPORTPRI!PGN = rsMeister!PGN
            Else
                rsIMPORTPRI!PGN = 0
            End If
            
            If Not IsNull(rsMeister!RKZ) Then
                rsIMPORTPRI!RKZ = rsMeister!RKZ
            Else
                rsIMPORTPRI!RKZ = "N"
            End If
            
            If rsMeister!MWST = "1" Then
                rsIMPORTPRI!MWST = "V"
            ElseIf rsMeister!MWST = "2" Then
                rsIMPORTPRI!MWST = "E"
            ElseIf rsMeister!MWST = "" Then
                rsIMPORTPRI!MWST = "V"
            ElseIf IsNull(rsMeister!MWST) Then
                rsIMPORTPRI!MWST = "V"
            Else
                rsIMPORTPRI!MWST = rsMeister!MWST
            End If
            
            rsIMPORTPRI!linr = rsMeister!linr
            rsIMPORTPRI!LIBESNR = rsMeister!LIBESNR
            
            If Not IsNull(rsMeister!EAN) Then
                sEAN = rsMeister!EAN
            Else
                sEAN = "0"
            End If
            
            If sEAN = "" Then
                rsIMPORTPRI!EAN = "0"
            Else
                rsIMPORTPRI!EAN = sEAN
            End If
            
            rsIMPORTPRI!EAN2 = rsMeister!EAN2
            rsIMPORTPRI!EAN3 = rsMeister!EAN3
            rsIMPORTPRI!ETIMERK = rsMeister!ETIMERK
            rsIMPORTPRI!MOPREIS = rsMeister!MOPREIS
            
            rsIMPORTPRI!vkpr = rsMeister!vkpr
            rsIMPORTPRI!KVKNEU = rsMeister!vkpr
            
            rsIMPORTPRI!lekpr = rsMeister!lekpr
            rsIMPORTPRI!NOTIZEN = rsMeister!NOTIZEN
            rsIMPORTPRI!BESTAND = rsMeister!BESTAND
            rsIMPORTPRI!VKMENGE = rsMeister!VKMENGE
            rsIMPORTPRI!VKDATUM = rsMeister!VKDATUM
            rsIMPORTPRI!MINMEN = rsMeister!MINMEN
            rsIMPORTPRI!INHALT = rsMeister!INHALT
            rsIMPORTPRI!INHALTBEZ = rsMeister!INHALTBEZ
            rsIMPORTPRI!GRUNDPREIS = rsMeister!GRUNDPREIS
            rsIMPORTPRI!MINBEST = rsMeister!MINBEST
            rsIMPORTPRI!RABATT_OK = rsMeister!RABATT_OK
            rsIMPORTPRI!GEFUEHRT = rsMeister!GEFUEHRT
            rsIMPORTPRI!KVKPR1 = rsMeister!KVKPR1
            rsIMPORTPRI!ekpr = rsMeister!ekpr
            rsIMPORTPRI!PREISSCHU = rsMeister!PREISSCHU
            rsIMPORTPRI!BONUS_OK = rsMeister!BONUS_OK
            rsIMPORTPRI!UMS_OK = rsMeister!UMS_OK
            rsIMPORTPRI!AWM = rsMeister!AWM
            rsIMPORTPRI!LASTDATE = rsMeister!LASTDATE
            rsIMPORTPRI!LASTTIME = rsMeister!LASTTIME
            rsIMPORTPRI!MNOTIZEN = rsMeister!Status
            rsIMPORTPRI!KVKalt = 0
            
            If Not IsNull(rsMeister!AUFDAT) Then
                rsIMPORTPRI!AUFDAT = rsMeister!AUFDAT
            Else
                rsIMPORTPRI!AUFDAT = Null
            End If
            
            If Not IsNull(rsMeister!EXDAT) Then
                rsIMPORTPRI!EXDAT = DateValue(Right(rsMeister!EXDAT, 2) & "." & Mid(rsMeister!EXDAT, 5, 2) & "." & Left(rsMeister!EXDAT, 4))
            Else
                rsIMPORTPRI!EXDAT = Null
            End If
            
            rsIMPORTPRI!FARBNR = rsMeister!FARBNR
            rsIMPORTPRI!MARKE = rsMeister!MARKE
            rsIMPORTPRI!GROESSE = rsMeister!GROESSE
            rsIMPORTPRI!SPANNE = rsMeister!SPANNE
            rsIMPORTPRI!AUFSCHLAG = rsMeister!AUFSCHLAG
            rsIMPORTPRI!SYNStatus = rsMeister!SYNStatus
           
            rsIMPORTPRI.Update
            rsMeister.MoveNext
        Loop
    End If
    rsMeister.Close
    rsIMPORTPRI.Close
    
    sSQL = "Delete from IMPORTPRI where EAN is null "
    gdApp.Execute sSQL, dbFailOnError
        Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "feldcheckMeister"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub GridFuellen()
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lWert       As Long
    Dim sWert       As String
    Dim dWert       As Double
    Dim siWert      As Single
    Dim rsrs        As Recordset
    Dim i           As Integer
    Dim j           As Integer
    Dim sSQL        As String
    Dim siAnzeige   As Single
    Dim lAnz        As Long
    
    sSQL = "Select * from IMPORTPRI order by awm , Linr,lpz"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    With PRIFlex
    .Redraw = False
    
    lrow = 1
    If Not rsrs.EOF Then
    
        rsrs.MoveLast
        
        Label2(35).Caption = rsrs.RecordCount
        Label2(35).Refresh
        
        lAnz = rsrs.RecordCount
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            siAnzeige = siAnzeige + 1
            txtStatus.Text = CStr((100 * siAnzeige) / lAnz)
            
            lrow = lrow + 1
            Label2(33).Caption = lrow
            Label2(33).Refresh
            
            .Rows = lrow + 1
            .Col = 0
            
            For i = 0 To byAnzahlSpalten - 1
                .Row = 0
                .Col = i
                
                If sSpaltenname(i) = .Text Then
                    
                    Select Case sSpaltenname(i)
                        Case Is = "Listen - EK", "Listen - VK", "KVK alt", "KVK neu"
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
                                
            If Not IsNull(rsrs!AWM) Then
                sWert = rsrs!AWM
                If Trim(sWert) = "98" Then 'neue Artikel
                    For j = 0 To byAnzahlSpalten - 1
                        .Col = j
                        .CellBackColor = vbWhite
                        .CellForeColor = &HFF&
                    Next j
                ElseIf Trim(sWert) = "97" Then
                    For j = 0 To byAnzahlSpalten - 1
                        .Col = j
                        .CellBackColor = vbYellow
                        .CellForeColor = vbBlue
                    Next j
                ElseIf Trim(sWert) = "96" Then
                    For j = 0 To byAnzahlSpalten - 1
                        .Col = j
                        .CellBackColor = vbBlue
                        .CellForeColor = vbWhite
                     Next j
                ElseIf Trim(sWert) = "95" Then
                    For j = 0 To byAnzahlSpalten - 1
                        .Col = j
                        .CellBackColor = vbBlue
                        .CellForeColor = &HFF&
                    Next j
                ElseIf Trim(sWert) = "94" Then 'doppelte EANs
                    For j = 0 To byAnzahlSpalten - 1
                        .Col = j
                        .CellBackColor = vbBlack
                        .CellForeColor = vbRed
                    
                    Next j
                End If
            End If
            rsrs.MoveNext
        Loop
    End If
    
    For i = 0 To byAnzahlSpalten - 1
        .Col = i
        .ColWidth(i) = aBreite(i) * 1.5
    Next i
        
    rsrs.Close
    picprogress.Visible = False
    Label2(35).Caption = ""
    Label2(35).Refresh
    Label2(33).Caption = ""
    Label2(33).Refresh
    If byAnzahlSpalten < 2 Then
    Else
        .FixedCols = 1
    End If
    .RowHeight(1) = 0
    lrow = lrow - 1
    If lrow = 0 Then
        anzeige "rot", "Es wurden keine Artikel ermittelt.", Label1(4)
    ElseIf lrow = 1 Then
        anzeige "Normal", "Ein Artikel wurde ermittelt.", Label1(4)
    Else
        anzeige "Normal", lrow & " Artikel wurden ermittelt.", Label1(4)
    End If
    .Redraw = True
    .Visible = True
    End With
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "GridFuellen"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub erstellegrid(cSpalte As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lrow        As Long
    Dim lcol        As Long
    Dim i           As Integer
    
    With MSFlexGrid1
        .Redraw = False
        .Visible = False
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .Row = 0
        
        .Col = 0
        .ColWidth(0) = 3000
        .Text = cSpalte
        
        .Col = 1
        .ColWidth(1) = 1500
        .Text = "Anz Artikel"
        
        .Col = 2
        .ColWidth(2) = 1500
        .Text = "auswählen"
        
        .Col = 3
        .ColWidth(3) = 2000
        .Text = ""
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "erstellegrid"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub zeigVorauswahlWerte(cDatname As String, cTabBlattname As String, cSpalte As String, bLK As Boolean)
On Error GoTo LOKAL_ERROR

    Dim gsExcel50   As String
    Dim dbExcel     As Database
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lWert       As Long
    Dim sWert       As String
    Dim lrow        As Long
    Dim i           As Integer
    Dim sLiefname   As String
    
    anzeige "normal", "", Label1(4)
    
    Screen.MousePointer = 11
    
    gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
        
    Set dbExcel = OpenDatabase(sExcelpfad & "\" & cDatname, 0, 0, gsExcel50)

    sSQL = " Select  [" & cSpalte & "],count([" & cSpalte & "]) as count from [" & cTabBlattname & "]"
    sSQL = sSQL & " where not [" & cSpalte & "] is null "
    sSQL = sSQL & " group by [" & cSpalte & "] "
    Set rsrs = dbExcel.OpenRecordset(sSQL)
    
    
    With MSFlexGrid1
    lrow = 1
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lrow = lrow + 1
            
            .Rows = lrow + 1
            .Row = lrow
            
            If Not IsNull(rsrs.Fields(0)) Then
                sWert = Trim(rsrs.Fields(0))
            Else
                sWert = ""
            End If
            
            If bLK = True Then
                sLiefname = ermBIOLK(sWert)
                
                .Col = 3
                .Text = sLiefname
            End If
            
            .Col = 0
            .Text = sWert
            
            If Not IsNull(rsrs.Fields(1)) Then
                lWert = Trim(rsrs.Fields(1))
            Else
                lWert = 0
            End If
            
            .Col = 1
            .Text = lWert
        
            .Col = 2
            .Text = "entfernt"
            .CellFontBold = True
            .CellForeColor = vbRed
                                   

            rsrs.MoveNext
        Loop
    End If
    
    If bLK = False Then
        .ColWidth(3) = 0
    Else
        .ColWidth(3) = 3000
    End If
    
    If UCase(cSpalte) = "LK" Then
        Check4.Visible = True
    Else
        Check4.Visible = False
    End If
    
   
    
    rsrs.Close
    dbExcel.Close
    
    .RowHeight(1) = 0
    lrow = lrow - 1
    
    .Redraw = True
    .Visible = True
    End With
    
    anzeige "normal", lrow & " Datensätze", Label2(32)
    
        
    Screen.MousePointer = 0
       
Exit Sub
LOKAL_ERROR:
    If err.Number = 3274 Or err.Number = 3170 Then
    
        Select Case gsExcel50
            Case Is = "Excel 3.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 4.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 4.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 5.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 5.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 7.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 7.0;HDR=yes;IMEX=1;"
                gsExcel50 = "Excel 8.0;HDR=yes;IMEX=1;"
                Resume
            Case Is = "Excel 8.0;HDR=yes;IMEX=1;"
                'Geht nicht noch unbekannt warum
                anzeige "rot1", "Diese Datei hat nicht das erwartete Format.", Label1(4)
        End Select
    ElseIf err.Number = 3349 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = Me.name
        Fehler.gsFunktion = "zeigVorauswahlWerte"
        Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Function ermBIOLK(sLiefkurz As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    ermBIOLK = sLiefkurz
    
    cSQL = "Select LIEFNAME from BIOPURLK where LIEFKURZ = '" & sLiefkurz & "'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LIEFNAME) Then
            ermBIOLK = rsrs!LIEFNAME
        End If
    End If
    rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "ermBIOLK"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub PricatStep4()
    On Error GoTo LOKAL_ERROR
    
    Dim j           As Integer
    Dim k           As Integer
    Dim cSQL        As String

    'Grid formatieren
    Tabcheck "STADAPRI"
    FormatGridOverTablay "STADAPRI"
    
    picprogress.Visible = True

    With PRIFlex
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
    
    'Daten ermitteln
    
    If Not FormatiereBildschirmdaten Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    'Grid fuellen
    anzeige "normal", "Die Daten werden angezeigt...", Label1(4)
    GridFuellen
    
    ermittlespalten
    Tabellenbreiteanpassen PRIFlex, 1.25 * gdTabfak
    
    For k = 0 To 11
        Check1(k).Visible = True
    Next
    
    If NewTableSuchenDBKombi("EXCELE", gdBase) Then
    
         If SpalteInTabellegefundenNEW("EXCELE", "transf9", gdBase) = False Then
            cSQL = " Alter table EXCELE add transf9 BIT  "
            gdBase.Execute cSQL, dbFailOnError
        End If
    
    
        lastvoreinstellungzeigen "EXCELE", frmWKL166, 9
    End If
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "PricatStep4"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
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
Private Sub File1_Click()
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cFileName As String

    If File1.ListIndex < 0 Then
    
    Else
        cFileName = File1.list(File1.ListIndex)
        zeigTabellenblätter cFileName, List1
        
        anzeige "WARN", cFileName, Label2(17)
        anzeige "WARN", cFileName, Label2(5)
        
        anzeige "WARN", cFileName, Label2(27)
    End If
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "File1_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List1_Click()
On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim cTabBlattname As String

    If List1.ListIndex < 0 Then
    
    Else
        cTabBlattname = List1.list(List1.ListIndex)
        zeigSpalten Label2(17), cTabBlattname, List2
        anzeige "WARN", cTabBlattname, Label2(8)
        anzeige "WARN", cTabBlattname, Label2(30)
    End If
       
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List1_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub List2_Click()
On Error GoTo LOKAL_ERROR
    Dim bFound As Boolean
    If List2.ListIndex < 0 Then

    Else
        zeigWerte True, Label2(5), Label2(8), List2.list(List2.ListIndex), List6, List7, "asc"
        If Check2.Value = vbUnchecked Then
            Check2.Caption = List2.list(List2.ListIndex)
        End If
        
        bFound = False
        If Label3(0).Caption = List2.list(List2.ListIndex) Then
            bFound = True
        End If
        
        If Label3(1).Caption = List2.list(List2.ListIndex) Then
            bFound = True
        End If
        
        If Label3(2).Caption = List2.list(List2.ListIndex) Then
            bFound = True
        End If
        
        
        If bFound Then
            Label3(0).Caption = ""
            Label3(1).Caption = ""
            Label3(2).Caption = ""
        End If
        
        If Label3(0).Caption = "" Then
            Label3(0).Caption = List2.list(List2.ListIndex)
        Else
            If Label3(1).Caption = "" Then
                If Label3(0).Caption <> List2.list(List2.ListIndex) Then
                    Label3(1).Caption = List2.list(List2.ListIndex)
                End If
            Else
                If Label3(2).Caption = "" Then
                    If Label3(1).Caption <> List2.list(List2.ListIndex) Then
                        Label3(2).Caption = List2.list(List2.ListIndex)
                    End If
                Else
                    Label3(0).Caption = List2.list(List2.ListIndex)
                    Label3(1).Caption = ""
                    Label3(2).Caption = ""
                End If
            End If
        End If
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List2_Click"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub List6_GotFocus()
On Error GoTo LOKAL_ERROR

    WerteGroßAnzeigen

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "List6_GotFocus"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub MSFlexGrid1_DblClick()
On Error GoTo LOKAL_ERROR

    Dim lcount As Long

    If MSFlexGrid1.Row = 1 Then
        sortierenGrid MSFlexGrid1
    Else
        If MSFlexGrid1.Col = 2 Then
    
            Select Case MSFlexGrid1.Text()
                Case "entfernt"
                    MSFlexGrid1.Text = "ausgewählt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbGreen
                    Command5(18).Enabled = True
                    Command5(18).BackColor = vbRed
                    
                Case "ausgewählt"
                    MSFlexGrid1.Text = "entfernt"
                    MSFlexGrid1.CellFontBold = True
                    MSFlexGrid1.CellForeColor = vbRed
                    
                    Command5(18).Enabled = False
                    Command5(18).BackColor = Command5(17).BackColor
                    
                    MSFlexGrid1.Redraw = False
                    For lcount = 1 To MSFlexGrid1.Rows - 1
                        MSFlexGrid1.Col = 2
                        MSFlexGrid1.Row = lcount
                
                        If MSFlexGrid1.Text = "ausgewählt" Then
                            Command5(18).Enabled = True
                            Command5(18).BackColor = vbRed
                            Exit For
                        End If
                    Next lcount
                    MSFlexGrid1.Redraw = True
                    
            End Select
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = Me.name
    Fehler.gsFunktion = "MSFlexGrid1_DblClick"
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo LOKAL_ERROR

If KeyCode = vbKeyF2 Then
    gF2Prompt.cFeld = ""
    gF2Prompt.cWert = ""
    gF2Prompt.cWert2 = ""
    gF2Prompt.cWahl = ""
    gF2Prompt.bMultiple = False
    
    Select Case Index
        Case 4
            gF2Prompt.cFeld = "AGN"
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
                Label1(5).Caption = ermAGNbez1(CLng(Val(Text1(Index).Text)))
            End If
        Case 2
            gF2Prompt.cFeld = "LINR"
            frmWK00a.Show 1
            If gF2Prompt.cWahl <> "" Then
                Text1(Index).Text = gF2Prompt.cWahl
                Label1(3).Caption = gF2Prompt.cWert
                
                Text1(0).Text = UCase(Trim(Left(gF2Prompt.cWert, 3)))
                
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
    Fehler.gsFehlertext = "Im Programmteil Excel Import ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
